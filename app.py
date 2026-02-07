import pandas as pd
import json
import io
import requests
import time 
import os
import sys
import tempfile
import traceback
import secrets
from flask import Flask, render_template, request, jsonify, send_file
import webbrowser
from threading import Timer
from datetime import datetime, timedelta
from openpyxl import Workbook
import numpy as np
from sklearn.cluster import DBSCAN
from math import radians
from sklearn.metrics.pairwise import haversine_distances
from math import radians, sin, cos, sqrt, atan2
import math

def resource_path(relative_path):
    """ PyInstaller için dosya yolu bulucu """
    try:
        # PyInstaller geçici klasörü
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

app = Flask(__name__, 
            template_folder=resource_path('templates'), 
            static_folder=resource_path('static'))

# Büyük projelerde "413 Request Entity Too Large" hatasını önlemek için istek gövdesi limiti (100 MB)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024

# EXE (PyInstaller) içinde tek instance UI kapanınca server'ı kapatmak için
APP_INSTANCE_TOKEN = secrets.token_urlsafe(16)
LAST_CLIENT_PING_TS = time.time()

@app.context_processor
def inject_runtime_flags():
    return {
        "shutdown_token": APP_INSTANCE_TOKEN,
        "is_frozen": bool(getattr(sys, "frozen", False)),
    }

@app.route('/ping', methods=['POST'])
def ping():
    """EXE UI açıkken heartbeat. UI kapanınca ping durur; watcher server'ı kapatır."""
    global LAST_CLIENT_PING_TS
    LAST_CLIENT_PING_TS = time.time()
    return jsonify({"status": "ok"})

@app.route('/shutdown', methods=['POST'])
def shutdown():
    """UI tarafından tetiklenebilir kontrollü kapatma (token kontrollü)."""
    data = request.get_json(silent=True) or {}
    if data.get("token") != APP_INSTANCE_TOKEN:
        return jsonify({"status": "error", "message": "unauthorized"}), 403

    shutdown_func = request.environ.get('werkzeug.server.shutdown')

    def _do_shutdown():
        try:
            if shutdown_func:
                shutdown_func()
            else:
                os._exit(0)
        except Exception:
            os._exit(0)

    # Response döndükten sonra kapat
    Timer(0.5, _do_shutdown).start()
    return jsonify({"status": "success"})

# --- LOJİSTİK PLANLAYICI SINIFI (ENTEGRE) ---
class LojistikPlanlayici:
    def __init__(self, osrm_url=None):
        if osrm_url is None:
            osrm_url = os.environ.get('OSRM_URL')
            if not osrm_url and os.environ.get('RENDER'):
                osrm_url = "http://router.project-osrm.org/route/v1/driving/"
            if not osrm_url:
                osrm_url = "http://localhost:5000/route/v1/driving/"
        u = str(osrm_url).strip()
        if not (u.startswith("http://") or u.startswith("https://")):
            u = "https://" + u
        if "route/v1/driving" not in u:
            if not u.endswith("/"):
                u += "/"
            u += "route/v1/driving/"
        else:
            if not u.endswith("/"):
                u += "/"
        self.osrm_url = u
        self.max_daily_minutes = 450  # 7.5 Saat (Rota optimizasyon sistemi ile aynı)
        
    def get_distance_duration(self, c1, c2):
        try:
            url = f"{self.osrm_url}{c1[1]},{c1[0]};{c2[1]},{c2[0]}?overview=false"
            # Render veya Public sunucu yavaş olabilir, timeout süresini artırıyoruz (0.5 -> 5.0 sn)
            r = requests.get(url, timeout=5.0).json()
            if r['code'] == 'Ok':
                return r['routes'][0]['distance']/1000, r['routes'][0]['duration']/60
        except: pass
        dist = self.haversine(c1, c2) * 1.3
        return dist, dist * 1.2

    def haversine(self, c1, c2):
        R = 6371
        dlat = radians(c2[0] - c1[0])
        dlon = radians(c2[1] - c1[1])
        a = math.sin(dlat/2)**2 + math.cos(radians(c1[0])) * math.cos(radians(c2[0])) * math.sin(dlon/2)**2
        return R * 2 * math.atan2(math.sqrt(a), math.sqrt(1-a))

    def split_into_days(self, sorted_stores, home_coords):
        days = []
        current_day = []
        
        if not sorted_stores: return days
        
        for store in sorted_stores:
            # Mağaza süresini güvenli şekilde al
            service_time = 30  # Varsayılan
            try:
                service_time_val = store.get('Mağaza İçin Süre', 30)
                # numpy/pandas tiplerini Python native'e çevir
                if hasattr(service_time_val, 'item'):
                    service_time = float(service_time_val.item())
                elif isinstance(service_time_val, (int, float)):
                    service_time = float(service_time_val)
                else:
                    service_time = float(service_time_val)
            except (ValueError, TypeError, AttributeError):
                service_time = 30
            
            # Negatif veya geçersiz değerleri varsayılan yap
            if not isinstance(service_time, (int, float)) or service_time <= 0:
                service_time = 30
            
            # Mevcut gündeki toplam mağaza süresini hesapla (SADECE Mağaza İçin Süre)
            current_day_service_time = 0.0
            for s in current_day:
                try:
                    s_time_val = s.get('Mağaza İçin Süre', 0)
                    if hasattr(s_time_val, 'item'):
                        s_time = float(s_time_val.item())
                    elif isinstance(s_time_val, (int, float)):
                        s_time = float(s_time_val)
                    else:
                        s_time = float(s_time_val)
                    if isinstance(s_time, (int, float)) and s_time > 0:
                        current_day_service_time += s_time
                except (ValueError, TypeError, AttributeError):
                    pass
            
            # 450 dakika limitini kontrol et (SADECE mağaza süreleri toplamı)
            if current_day_service_time + service_time > self.max_daily_minutes:
                # Limit aşılıyor, yeni güne geç
                if current_day:
                    days.append(current_day)
                current_day = [store]
            else:
                # Limit içinde, mevcut güne ekle
                current_day.append(store)
        
        if current_day:
            days.append(current_day)
            
        return days

    def _is_sunday_day_number(self, day_number: int) -> bool:
        """
        Gün bazlı planlamada haftanın günleri varsayımı:
        1 = Pazartesi, 7 = Pazar (yani 7, 14, 21, 28 ... pazar)
        """
        try:
            dn = int(day_number)
        except Exception:
            return False
        return dn % 7 == 0

    def _normalize_day_in_cycle(self, day_number: int, cycle_days: int = 28) -> int:
        """Günü 1..cycle_days aralığına sar (wrap)."""
        try:
            dn = int(day_number)
        except Exception:
            dn = 1
        if cycle_days <= 0:
            return dn
        # Python mod negatifleri de destekler, güvenli normalize
        return ((dn - 1) % cycle_days) + 1

    def process(self, df, start_date_str=None, start_day=None, cycle_days: int = 28):
        final_rows = []
        
        # Veri temizliği
        if 'Personel' not in df.columns or df['Personel'].isnull().all():
             df['Personel'] = 'Personel 1'
        
        df['Personel'] = df['Personel'].fillna('Atanmamış')
        
        # GÜNCELLEME: Mağaza İçin Süre sütununu numeric'e çevir
        if 'Mağaza İçin Süre' in df.columns:
            df['Mağaza İçin Süre'] = pd.to_numeric(df['Mağaza İçin Süre'], errors='coerce')
            df['Mağaza İçin Süre'] = df['Mağaza İçin Süre'].fillna(30)  # Varsayılan 30 dakika
        else:
            df['Mağaza İçin Süre'] = 30  # Sütun yoksa varsayılan değer

        # Merkez Koordinat (Tüm verinin ortalaması)
        center_lat = df['Enlem'].mean()
        center_lon = df['Boylam'].mean()
        home_coords = (center_lat, center_lon)

        for p_name, p_df in df.groupby('Personel'):
            # Mağaza İçin Süre'yi Python native float'a çevir
            p_df = p_df.copy()
            if 'Mağaza İçin Süre' in p_df.columns:
                p_df['Mağaza İçin Süre'] = p_df['Mağaza İçin Süre'].astype(float)
            stores_list = p_df.to_dict('records')
            
            # 1. En Yakın Komşu Sıralaması (Rota Oluşturma)
            unvisited = stores_list.copy()
            path = []
            curr = home_coords
            
            while unvisited:
                # En yakını bul
                next_store = min(unvisited, key=lambda x: self.haversine(curr, (x['Enlem'], x['Boylam'])))
                path.append(next_store)
                unvisited.remove(next_store)
                curr = (next_store['Enlem'], next_store['Boylam'])
            
            # 2. Günlere Bölme
            days = self.split_into_days(path, home_coords)
            
            # 3. Plan Atama (Tarih bazlı veya Gün bazlı)
            use_day_mode = start_day is not None

            if use_day_mode:
                try:
                    current_day = int(start_day) if str(start_day).strip() != "" else 1
                except Exception:
                    current_day = 1

                # 1..cycle_days aralığında başlat
                current_day = self._normalize_day_in_cycle(current_day, cycle_days=cycle_days)

                for day_stores in days:
                    # Pazar gününü atla (7, 14, 21, 28 ...)
                    guard = 0
                    while self._is_sunday_day_number(current_day):
                        current_day += 1
                        current_day = self._normalize_day_in_cycle(current_day, cycle_days=cycle_days)
                        guard += 1
                        if guard > (cycle_days + 7):
                            # Güvenlik: sonsuz döngü olmasın
                            break

                    day_label = f"{current_day}. Gün"

                    for store in day_stores:
                        row = store.copy()
                        row['Atanan Tarih'] = day_label
                        final_rows.append(row)

                    current_day += 1
                    current_day = self._normalize_day_in_cycle(current_day, cycle_days=cycle_days)
            else:
                try:
                    current_date = datetime.strptime(start_date_str, '%Y-%m-%d')
                except Exception:
                    current_date = datetime.now()

                for day_stores in days:
                    # Pazar gününü atla
                    while current_date.weekday() == 6:
                        current_date += timedelta(days=1)
                    
                    date_str = current_date.strftime('%d.%m.%Y') # Ön yüz formatı: GG.AA.YYYY
                    
                    for store in day_stores:
                        row = store.copy()
                        row['Atanan Tarih'] = date_str
                        final_rows.append(row)
                    
                    current_date += timedelta(days=1)

        return pd.DataFrame(final_rows)

LAST_UPLOADED_FILE_PATH = None
OSRM_API_BASE_URL = "http://localhost:5000/route/v1/driving/"

class RotaOptimizasyonSistemi:
    def __init__(self, osrm_url="http://localhost:5000/route/v1/driving/"):
        self.osrm_url = osrm_url
        self.max_daily_minutes = 450
        
    def get_osrm_route_info(self, coord1, coord2):
        """OSRM ile iki nokta arası mesafe ve süreyi hesapla"""
        try:
            lon1, lat1 = coord1[1], coord1[0]
            lon2, lat2 = coord2[1], coord2[0]
            
            url = f"{self.osrm_url}{lon1},{lat1};{lon2},{lat2}?overview=false"
            response = requests.get(url, timeout=5)
            data = response.json()
            
            if data['code'] == 'Ok' and len(data['routes']) > 0:
                route = data['routes'][0]
                distance_km = route['distance'] / 1000  # metre -> km
                duration_min = route['duration'] / 60   # saniye -> dakika
                return distance_km, duration_min
            else:
                return self.haversine_distance(coord1, coord2), self.estimate_travel_time(coord1, coord2)
        except Exception as e:
            print(f"OSRM hatası: {e}")
            dist = self.haversine_distance(coord1, coord2)
            time_val = self.estimate_travel_time(coord1, coord2)
            return dist, time_val

    def haversine_distance(self, coord1, coord2):
        """Haversine formülü ile mesafe hesaplama"""
        lat1, lon1 = coord1
        lat2, lon2 = coord2
        
        lat1_rad = radians(lat1)
        lon1_rad = radians(lon1)
        lat2_rad = radians(lat2)
        lon2_rad = radians(lon2)
        
        dlon = lon2_rad - lon1_rad
        dlat = lat2_rad - lat1_rad
        a = math.sin(dlat/2)**2 + math.cos(lat1_rad) * math.cos(lat2_rad) * math.sin(dlon/2)**2
        c = 2 * math.atan2(math.sqrt(a), math.sqrt(1-a))
        radius_earth = 6371
        return radius_earth * c

    def estimate_travel_time(self, coord1, coord2):
        """Tahmini seyahat süresi hesaplama"""
        distance = self.haversine_distance(coord1, coord2)
        return distance * 2

    def calculate_distance_matrix(self, coordinates):
        """Tüm koordinatlar arası mesafe matrisi hesapla"""
        n = len(coordinates)
        distance_matrix = np.zeros((n, n))
        time_matrix = np.zeros((n, n))
        
        print(f"Mesafe matrisi hesaplanıyor... {n}x{n}")
        
        for i in range(n):
            for j in range(n):
                if i != j:
                    dist, time_val = self.get_osrm_route_info(coordinates[i], coordinates[j])
                    distance_matrix[i][j] = dist
                    time_matrix[i][j] = time_val
                else:
                    distance_matrix[i][j] = 0
                    time_matrix[i][j] = 0
                    
        return distance_matrix, time_matrix

    def dbscan_clustering(self, coordinates, max_radius_km=1, min_stores_per_cluster=1):
        """DBSCAN ile mağazaları coğrafi olarak kümele"""
        epsilon = max_radius_km * 0.00899321606
        
        dbscan = DBSCAN(eps=epsilon, min_samples=min_stores_per_cluster, metric='haversine')
        clusters = dbscan.fit_predict(np.radians(coordinates))
        
        noise_mask = clusters == -1
        if noise_mask.any():
            print(f"Küme atanamayan {noise_mask.sum()} mağaza var, en yakın kümeye atanıyor...")
            
            for i in np.where(noise_mask)[0]:
                min_dist = float('inf')
                closest_cluster = -1
                
                for j in range(len(coordinates)):
                    if clusters[j] != -1:
                        dist = self.haversine_distance(coordinates[i], coordinates[j])
                        if dist < min_dist and dist <= 5.0:
                            min_dist = dist
                            closest_cluster = clusters[j]
                
                if closest_cluster != -1:
                    clusters[i] = closest_cluster
                    print(f"  Mağaza {i} {min_dist:.2f} km uzaklıktaki kümeye atandı")
                else:
                    new_cluster = np.max(clusters) + 1 if len(np.unique(clusters)) > 1 else 0
                    clusters[i] = new_cluster
                    print(f"  Mağaza {i} yeni kümeye atandı (çok uzak)")
        
        return clusters

    def adaptive_dbscan_clustering(self, coordinates, personel_sayisi):
        """Personel sayısına göre otomatik yarıçap ayarlayan DBSCAN"""
        max_radius_options = [1, 2, 3, 5, 8, 10, 15]
        
        best_clusters = None
        best_radius = None
        
        for max_radius in max_radius_options:
            print(f"{max_radius} km yarıçap ile kümeleme deneniyor...")
            clusters = self.dbscan_clustering(coordinates, max_radius, min_stores_per_cluster=1)
            unique_clusters = len(np.unique(clusters))
            
            print(f"Oluşan küme sayısı: {unique_clusters}")
            
            if unique_clusters == personel_sayisi:
                return clusters
            elif best_clusters is None or abs(unique_clusters - personel_sayisi) < abs(len(np.unique(best_clusters)) - personel_sayisi):
                best_clusters = clusters
                best_radius = max_radius
        
        print(f"En iyi yarıçap: {best_radius} km, küme sayısı: {len(np.unique(best_clusters))}")
        
        if len(np.unique(best_clusters)) > personel_sayisi:
            best_clusters = self.merge_small_clusters(coordinates, best_clusters, personel_sayisi)
        elif len(np.unique(best_clusters)) < personel_sayisi:
            best_clusters = self.split_large_clusters(coordinates, best_clusters, personel_sayisi)
        
        return best_clusters

    def split_large_clusters(self, coordinates, clusters, target_cluster_count):
        """Büyük kümeleri daha küçük parçalara böl"""
        unique_clusters = np.unique(clusters)
        current_count = len(unique_clusters)
        
        while current_count < target_cluster_count:
            cluster_sizes = [np.sum(clusters == c) for c in unique_clusters]
            largest_cluster_idx = np.argmax(cluster_sizes)
            largest_cluster = unique_clusters[largest_cluster_idx]
            
            cluster_points = coordinates[clusters == largest_cluster]
            if len(cluster_points) >= 2:
                from sklearn.cluster import KMeans
                kmeans = KMeans(n_clusters=2, random_state=42, n_init=10)
                sub_clusters = kmeans.fit_predict(cluster_points)
                
                new_cluster_id = np.max(clusters) + 1
                cluster_indices = np.where(clusters == largest_cluster)[0]
                for i, sub_cluster in enumerate(sub_clusters):
                    if sub_cluster == 1:
                        clusters[cluster_indices[i]] = new_cluster_id
                
                current_count += 1
                unique_clusters = np.unique(clusters)
            else:
                break
        
        return clusters

    def merge_small_clusters(self, coordinates, clusters, target_cluster_count):
        """Küçük kümeleri birleştir"""
        unique_clusters = np.unique(clusters)
        cluster_sizes = [np.sum(clusters == c) for c in unique_clusters]
        
        while len(unique_clusters) > target_cluster_count:
            sorted_clusters = [c for _, c in sorted(zip(cluster_sizes, unique_clusters))]
            smallest1, smallest2 = sorted_clusters[:2]
            
            clusters[clusters == smallest2] = smallest1
            
            unique_clusters = np.unique(clusters)
            cluster_sizes = [np.sum(clusters == c) for c in unique_clusters]
        
        return clusters

    def nearest_neighbor_route(self, stores_df, distance_matrix, time_matrix):
        """En yakın komşu algoritması ile optimal rota oluştur"""
        n = len(stores_df)
        if n <= 1:
            return self.create_simple_route(stores_df, distance_matrix, time_matrix)
        
        print(f"En yakın komşu algoritması ile {n} mağaza için rota oluşturuluyor...")
        
        unvisited = set(range(n))
        route = []
        
        center_lat = stores_df['Enlem'].mean()
        center_lon = stores_df['Boylam'].mean()
        
        min_dist = float('inf')
        current = 0
        for i in range(n):
            dist = self.haversine_distance(
                (center_lat, center_lon), 
                (stores_df.iloc[i]['Enlem'], stores_df.iloc[i]['Boylam'])
            )
            if dist < min_dist:
                min_dist = dist
                current = i
        
        route.append(current)
        unvisited.remove(current)
        
        while unvisited:
            next_node = None
            min_distance = float('inf')
            
            for candidate in unvisited:
                if distance_matrix[current][candidate] < min_distance:
                    min_distance = distance_matrix[current][candidate]
                    next_node = candidate
            
            if next_node is not None:
                route.append(next_node)
                unvisited.remove(next_node)
                current = next_node
            else:
                break
        
        ordered_stores = []
        for i, store_idx in enumerate(route):
            store_info = stores_df.iloc[store_idx].copy()
            
            if i == 0:
                store_info['Bir önceki mağazaya Km'] = 0
                store_info['Bir önceki Mağazaya süre (dk)'] = 0
            else:
                prev_store_idx = route[i-1]
                dist = distance_matrix[prev_store_idx][store_idx]
                time_val = time_matrix[prev_store_idx][store_idx]
                store_info['Bir önceki mağazaya Km'] = round(dist, 2)
                store_info['Bir önceki Mağazaya süre (dk)'] = round(time_val, 2)
            
            ordered_stores.append(store_info)
        
        return ordered_stores

    def create_simple_route(self, stores_df, distance_matrix, time_matrix):
        """Basit rota oluştur"""
        stores = stores_df.copy()
        route = []
        
        stores = stores.sort_values(['Enlem', 'Boylam'])
        
        for i, (idx, store) in enumerate(stores.iterrows()):
            store_info = store.copy()
            
            if i == 0:
                store_info['Bir önceki mağazaya Km'] = 0
                store_info['Bir önceki Mağazaya süre (dk)'] = 0
            else:
                prev_idx = stores.index[i-1]
                prev_store_idx = stores.index.get_loc(prev_idx)
                current_store_idx = stores.index.get_loc(idx)
                
                dist = distance_matrix[prev_store_idx][current_store_idx]
                time_val = time_matrix[prev_store_idx][current_store_idx]
                
                store_info['Bir önceki mağazaya Km'] = round(dist, 2)
                store_info['Bir önceki Mağazaya süre (dk)'] = round(time_val, 2)
            
            route.append(store_info)
        
        return route

    def solve_vrp_for_cluster(self, stores_df, distance_matrix, time_matrix):
        """En yakın komşu algoritması ile rota oluştur"""
        return self.nearest_neighbor_route(stores_df, distance_matrix, time_matrix)

    def optimize_daily_routes(self, personel_stores, distance_matrix, time_matrix):
        """Bir personelin mağazalarını günlere böl"""
        if len(personel_stores) == 0:
            return []
            
        ordered_stores = self.nearest_neighbor_route(personel_stores, distance_matrix, time_matrix)
        
        daily_routes = []
        current_day = 1
        current_day_stores = []
        current_total_time = 0
        
        for i, store in enumerate(ordered_stores):
            store_time = store['Mağaza İçin Süre']
            travel_time = store.get('Bir önceki Mağazaya süre (dk)', 0)
            
            total_time_for_store = travel_time + store_time
            
            if current_total_time + total_time_for_store <= self.max_daily_minutes:
                current_day_stores.append(store)
                current_total_time += total_time_for_store
            else:
                if current_day_stores:
                    for store_info in current_day_stores:
                        store_info['Plan Tarihi'] = f'{current_day}. Gün'
                    daily_routes.extend(current_day_stores)
                    current_day += 1
                    current_day_stores = [store]
                    current_total_time = store_time
                else:
                    store['Plan Tarihi'] = f'{current_day}. Gün'
                    daily_routes.append(store)
                    current_day += 1
        
        if current_day_stores:
            for store_info in current_day_stores:
                store_info['Plan Tarihi'] = f'{current_day}. Gün'
            daily_routes.extend(current_day_stores)
        
        return daily_routes

    def optimize_routes(self, stores_df, personel_sayisi):
        """Ana optimizasyon fonksiyonu"""
        print("Optimizasyon başlıyor...")
        print(f"Toplam mağaza sayısı: {len(stores_df)}")
        print(f"Personel sayısı: {personel_sayisi}")
        
        coordinates = stores_df[['Enlem', 'Boylam']].values
        
        print("DBSCAN ile kümeleme yapılıyor...")
        clusters = self.adaptive_dbscan_clustering(coordinates, personel_sayisi)
        
        unique_clusters = np.unique(clusters)
        print(f"Oluşan küme sayısı: {len(unique_clusters)}")
        
        stores_df['Personel'] = [f'Personel {int(cluster)+1}' for cluster in clusters]
        
        print("Mesafe matrisi hesaplanıyor...")
        distance_matrix, time_matrix = self.calculate_distance_matrix(coordinates)
        
        final_routes = []
        
        for personel_id in range(len(unique_clusters)):
            personel_name = f'Personel {personel_id+1}'
            print(f"{personel_name} için rota planlanıyor...")
            
            personel_stores = stores_df[stores_df['Personel'] == personel_name].copy()
            print(f"{personel_name} mağaza sayısı: {len(personel_stores)}")
            
            if len(personel_stores) > 0:
                personel_indices = personel_stores.index.tolist()
                personel_coordinates = personel_stores[['Enlem', 'Boylam']].values
                
                personel_dist_matrix, personel_time_matrix = self.calculate_distance_matrix(personel_coordinates)
                
                daily_routes = self.optimize_daily_routes(personel_stores, personel_dist_matrix, personel_time_matrix)
                final_routes.extend(daily_routes)
        
        result_df = pd.DataFrame(final_routes)
        
        original_columns = ['Mağaza İsmi', 'Mağaza İçin Süre', 'Enlem', 'Boylam']
        new_columns = ['Personel', 'Plan Tarihi', 'Bir önceki mağazaya Km', 'Bir önceki Mağazaya süre (dk)']
        all_columns = original_columns + new_columns
        
        for col in all_columns:
            if col not in result_df.columns:
                result_df[col] = 0 if 'Km' in col or 'süre' in col else ''
        
        print(f"Optimizasyon tamamlandı. Toplam {len(result_df)} mağaza dağıtıldı.")
        
        self.print_statistics(result_df)
        
        return result_df[all_columns]

    def print_statistics(self, result_df):
        """İstatistikleri yazdır"""
        print("\n=== OPTİMİZASYON İSTATİSTİKLERİ ===")
        
        for personel in result_df['Personel'].unique():
            personel_df = result_df[result_df['Personel'] == personel]
            gun_sayisi = personel_df['Plan Tarihi'].nunique()
            
            print(f"\n{personel}:")
            print(f"  Toplam mağaza: {len(personel_df)}")
            print(f"  Toplam gün: {gun_sayisi}")
            
            for gun in sorted(personel_df['Plan Tarihi'].unique()):
                gun_df = personel_df[personel_df['Plan Tarihi'] == gun]
                toplam_ziyaret = gun_df['Mağaza İçin Süre'].sum()
                toplam_seyahat = gun_df['Bir önceki Mağazaya süre (dk)'].sum()
                toplam_sure = toplam_ziyaret + toplam_seyahat
                
                print(f"  {gun}: {len(gun_df)} mağaza, {toplam_sure:.0f} dakika "
                      f"(Ziyaret: {toplam_ziyaret}dk, Seyahat: {toplam_seyahat:.0f}dk)")

def open_browser():
    webbrowser.open_new("http://127.0.0.1:5001")

def _haversine_km(lat1, lon1, lat2, lon2):
    """Kuş uçuşu mesafe (km)."""
    R = 6371.0
    dlat = radians(lat2 - lat1)
    dlon = radians(lon2 - lon1)
    a = sin(dlat/2)**2 + cos(radians(lat1)) * cos(radians(lat2)) * sin(dlon/2)**2
    return R * 2 * atan2(sqrt(a), sqrt(1-a))

def get_route_details(lat1, lon1, lat2, lon2):
    coordinates = f"{lon1},{lat1};{lon2},{lat2}"
    url = OSRM_API_BASE_URL + coordinates
    for attempt in range(3):
        try:
            response = requests.get(url, timeout=20)
            response.raise_for_status()
            data = response.json()
            if data.get('code') == 'Ok' and data.get('routes'):
                distance_m = data['routes'][0]['distance']
                duration_s = data['routes'][0]['duration']
                distance_km = round(distance_m / 1000, 2)
                duration_min = round(duration_s / 60, 2)
                return distance_km, duration_min
        except requests.exceptions.RequestException as e:
            print(f"Rota servisi hatası (Deneme {attempt + 1}/3): {e}")
            time.sleep(1)

    # OSRM yoksa kuş uçuşu fallback
    try:
        km = _haversine_km(float(lat1), float(lon1), float(lat2), float(lon2))
        # Basit süre tahmini: 1 km ~ 2 dk (uygulamadaki diğer tahmin mantığıyla uyumlu)
        min_val = km * 2
        return round(km, 2), round(min_val, 2)
    except Exception:
        return 0, 0

def convert_to_serializable(obj):
    """Tüm veri tiplerini JSON serializable hale getir"""
    if pd.isna(obj) or obj is None:
        return ""
    elif hasattr(obj, 'strftime'):  # Timestamp, datetime vb.
        try:
            return obj.strftime('%d.%m.%Y')
        except:
            return str(obj)
    elif isinstance(obj, (int, float)):
        return obj
    else:
        return str(obj)

@app.route('/', methods=['GET', 'POST'])
def index():
    global LAST_UPLOADED_FILE_PATH
    store_data = []
    initial_schedule = {}
    all_merch = set()
    
    try:
        if request.method == 'POST':
            if 'file' not in request.files: 
                return render_template('index.html', error="Dosya seçilmedi.", store_data="[]", initial_schedule="{}", all_merch=[])
            
            file = request.files.get('file')
            if not file or file.filename == '': 
                return render_template('index.html', error="Dosya seçilmedi.", store_data="[]", initial_schedule="{}", all_merch=[])
            
            if file and (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
                try:
                    # Excel'i oku - tüm sütunları string olarak oku
                    df = pd.read_excel(file, dtype=str)
                    
                    # Sütun isimlerini yeni formata çevir
                    column_mapping = {
                        'Store': 'Mağaza İsmi',
                        'Lat': 'Enlem', 
                        'Long': 'Boylam',
                        'Merchandiser': 'Personel',
                        'Duration': 'Mağaza İçin Süre',
                        'Gün': 'Plan Tarihi',
                        'Şehir': 'İl'
                    }
                    
                    # Eski sütun isimlerini yeni isimlere çevir
                    df.rename(columns=column_mapping, inplace=True)
                    
                    # Gerekli sütunları kontrol et (yeni isimlerle)
                    required_cols = ['Mağaza İsmi', 'Enlem', 'Boylam']
                    if not all(col in df.columns for col in required_cols):
                        return render_template('index.html', error="Excel'de 'Mağaza İsmi', 'Enlem', 'Boylam' sütunları bulunmalıdır.", store_data="[]", initial_schedule="{}", all_merch=[])
                    
                    # Koordinatları numeric'e çevir (harita için gerekli)
                    try:
                        df['Enlem'] = pd.to_numeric(df['Enlem'], errors='coerce')
                        df['Boylam'] = pd.to_numeric(df['Boylam'], errors='coerce')
                    except Exception as e:
                        return render_template('index.html', error=f"Koordinatlar numeric değere dönüştürülemedi: {str(e)}", store_data="[]", initial_schedule="{}", all_merch=[])
                    
                    # Mağaza İçin Süre sütununu numeric'e çevir
                    if 'Mağaza İçin Süre' in df.columns:
                        try:
                            df['Mağaza İçin Süre'] = pd.to_numeric(df['Mağaza İçin Süre'], errors='coerce')
                        except:
                            pass
                    
                    # Geçersiz koordinatları filtrele
                    df = df.dropna(subset=['Enlem', 'Boylam'])
                    
                    if df.empty:
                        return render_template('index.html', error="Geçerli koordinat bulunamadı.", store_data="[]", initial_schedule="{}", all_merch=[])
                    
                    # Tüm verileri temizle ve string'e çevir
                    for col in df.columns:
                        if df[col].dtype == 'object':
                            df[col] = df[col].fillna('')
                            try:
                                df[col] = df[col].apply(lambda x: convert_to_serializable(x))
                            except Exception as e:
                                print(f"Sütun {col} temizlenirken hata: {e}")
                                df[col] = df[col].astype(str)
                    
                    # Tüm sütunları koru, hiçbir sütunu silme
                    temp_dir = tempfile.gettempdir()
                    temp_path = os.path.join(temp_dir, "route_planner_temp_data.csv")
                    df.to_csv(temp_path, index=False, encoding='utf-8')
                    LAST_UPLOADED_FILE_PATH = temp_path

                    # Plan Tarihi sütunu varsa schedule verisini hazırla
                    if 'Plan Tarihi' in df.columns:
                        schedule_df = df[df['Plan Tarihi'].notna() & (df['Plan Tarihi'] != '')].copy()
                        if not schedule_df.empty:
                            initial_schedule = {}
                            for store_name, group in schedule_df.groupby('Mağaza İsmi'):
                                dates = group['Plan Tarihi'].tolist()
                                # Tarihleri temizle ve formatla
                                cleaned_dates = []
                                for date_str in dates:
                                    if date_str and str(date_str).strip() and str(date_str).strip().lower() not in ['nan', 'none', '']:
                                        cleaned_dates.append(str(date_str).strip())
                                if cleaned_dates:
                                    initial_schedule[store_name] = cleaned_dates
                    
                    # Store_data'yı oluştur - tüm veriler string
                    store_data = []
                    for _, row in df.iterrows():
                        store_dict = {}
                        for col in df.columns:
                            store_dict[col] = str(row[col]) if pd.notna(row[col]) else ""
                        store_data.append(store_dict)

                    # Tüm personelleri topla
                    for store in store_data:
                        try:
                            personel = store.get('Personel', '')
                            if personel:
                                # Personel değerini string'e çevir ve kontrol et
                                personel_str = str(personel).strip() if isinstance(personel, str) else str(personel)
                                if personel_str and personel_str.lower() not in ['nan', 'none', '']:
                                    all_merch.add(personel_str)
                        except Exception as e:
                            print(f"Personel işlenirken hata: {e}")
                            continue

                except Exception as e:
                    print(traceback.format_exc())
                    return render_template('index.html', error=f"Dosya okunurken hata: {str(e)}", store_data="[]", initial_schedule="{}", all_merch=[])
        
        # JSON'a çevir - tüm veriler string olduğu için sorun çıkmaz
        try:
            stores_json = json.dumps(store_data, ensure_ascii=False, default=str)
            schedule_json = json.dumps(initial_schedule, ensure_ascii=False, default=str)
            all_merch = sorted(list(all_merch))
        except Exception as e:
            print(f"JSON dönüşüm hatası: {e}")
            print(traceback.format_exc())
            stores_json = "[]"
            schedule_json = "{}"
            all_merch = []
        
        return render_template('index.html', store_data=stores_json, initial_schedule=schedule_json, all_merch=all_merch)
    
    except Exception as e:
        print(f"Index route genel hatası: {e}")
        print(traceback.format_exc())
        return render_template('index.html', error=f"Beklenmeyen bir hata oluştu: {str(e)}", store_data="[]", initial_schedule="{}", all_merch=[])

@app.route('/get_route', methods=['POST'])
def get_route():
    data = request.get_json()
    start = data['start']
    end = data['end']
    
    lat1, lon1 = start['lat'], start['lng']
    lat2, lon2 = end['lat'], end['lng']
    
    km, min_val = get_route_details(lat1, lon1, lat2, lon2)
    
    if km > 0 and min_val > 0:
        return jsonify({
            'status': 'success',
            'km': km,
            'min': min_val
        })
    else:
        return jsonify({
            'status': 'error',
            'message': 'Rota hesaplanamadı'
        })

@app.route('/calculate_multi_point_route', methods=['POST'])
def calculate_multi_point_route():
    data = request.get_json()
    points = data['points']
    
    if len(points) < 2:
        return jsonify({
            'status': 'error',
            'message': 'En az 2 nokta gereklidir'
        })
    
    try:
        # OSRM için koordinatları hazırla
        coordinates = []
        for point in points:
            coordinates.append(f"{point['lon']},{point['lat']}")
        
        coordinates_str = ";".join(coordinates)
        url = f"{OSRM_API_BASE_URL}{coordinates_str}?overview=full&geometries=geojson"
        
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        osrm_data = response.json()
        
        if osrm_data.get('code') == 'Ok' and osrm_data.get('routes'):
            route = osrm_data['routes'][0]
            distance_m = route['distance']
            duration_s = route['duration']
            
            # Geometriyi koordinat listesine çevir
            geometry = route['geometry']['coordinates']
            path = [[coord[1], coord[0]] for coord in geometry]
            
            return jsonify({
                'status': 'success',
                'km': round(distance_m / 1000, 2),
                'min': round(duration_s / 60, 2),
                'path': path
            })
        else:
            raise RuntimeError("OSRM rota hesaplayamadı")
            
    except requests.exceptions.RequestException as e:
        print(f"OSRM rota hatası: {e}")
        # Kuş uçuşu fallback
        path = [[float(p['lat']), float(p['lon'])] for p in points]
        km = 0.0
        for i in range(len(path) - 1):
            km += _haversine_km(path[i][0], path[i][1], path[i+1][0], path[i+1][1])
        min_val = km * 2
        return jsonify({
            'status': 'success',
            'km': round(km, 2),
            'min': round(min_val, 2),
            'path': path,
            'fallback': True,
            'message': 'OSRM bağlantısı yok, kuş uçuşu rota kullanıldı.'
        })
    except Exception as e:
        print(f"Genel rota hatası: {e}")
        # Kuş uçuşu fallback
        try:
            path = [[float(p['lat']), float(p['lon'])] for p in points]
            km = 0.0
            for i in range(len(path) - 1):
                km += _haversine_km(path[i][0], path[i][1], path[i+1][0], path[i+1][1])
            min_val = km * 2
            return jsonify({
                'status': 'success',
                'km': round(km, 2),
                'min': round(min_val, 2),
                'path': path,
                'fallback': True,
                'message': 'OSRM yok, kuş uçuşu rota kullanıldı.'
            })
        except Exception:
            return jsonify({
                'status': 'error',
                'message': f'Beklenmeyen hata: {str(e)}'
            })

# YENİ: Çoklu mağaza yükleme endpoint'i
@app.route('/upload_stores', methods=['POST'])
def upload_stores():
    try:
        if 'file' not in request.files:
            return jsonify({'status': 'error', 'message': 'Dosya seçilmedi'})
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'status': 'error', 'message': 'Dosya seçilmedi'})
        
        if file and (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
            # Excel'i oku
            df = pd.read_excel(file, dtype=str)
            
            # Gerekli sütunları kontrol et
            required_cols = ['Mağaza İsmi', 'Enlem', 'Boylam']
            if not all(col in df.columns for col in required_cols):
                return jsonify({'status': 'error', 'message': "Excel'de 'Mağaza İsmi', 'Enlem', 'Boylam' sütunları bulunmalıdır."})
            
            # Koordinatları numeric'e çevir
            try:
                df['Enlem'] = pd.to_numeric(df['Enlem'], errors='coerce')
                df['Boylam'] = pd.to_numeric(df['Boylam'], errors='coerce')
            except Exception as e:
                return jsonify({'status': 'error', 'message': f"Koordinatlar numeric değere dönüştürülemedi: {str(e)}"})
            
            # Geçersiz koordinatları filtrele
            df = df.dropna(subset=['Enlem', 'Boylam'])
            
            if df.empty:
                return jsonify({'status': 'error', 'message': "Geçerli koordinat bulunamadı."})
            
            # Diğer sütunları ekle (eğer yoksa)
            optional_cols = ['Personel', 'Mağaza İçin Süre', 'Plan Tarihi', 'İl', 'Bölge', 'Frekans', 'İlçe']
            for col in optional_cols:
                if col not in df.columns:
                    df[col] = ''
            
            # Verileri temizle
            for col in df.columns:
                if df[col].dtype == 'object':
                    df[col] = df[col].fillna('')
                    df[col] = df[col].apply(lambda x: convert_to_serializable(x))
            
            # DataFrame'i dictionary listesine çevir
            stores_list = []
            for _, row in df.iterrows():
                store_dict = {}
                for col in df.columns:
                    store_dict[col] = str(row[col]) if pd.notna(row[col]) else ""
                stores_list.append(store_dict)
            
            return jsonify({
                'status': 'success',
                'stores': stores_list,
                'count': len(stores_list)
            })
        
        else:
            return jsonify({'status': 'error', 'message': 'Geçersiz dosya formatı'})
            
    except Exception as e:
        print(f"Mağaza yükleme hatası: {e}")
        return jsonify({'status': 'error', 'message': f'Dosya işlenirken hata oluştu: {str(e)}'})

@app.route('/download', methods=['POST'])
def download():
    global LAST_UPLOADED_FILE_PATH
    
    # Proje adı verilmişse veriyi dosyadan yükle (büyük POST → 413 hatasını önler)
    project_name = request.form.get('project_name', '').strip()
    if project_name:
        safe_name = "".join(c for c in project_name if c.isalnum() or c in (' ', '-', '_')).strip()
        project_file = os.path.join(PROJECTS_DIR, f"{safe_name}.json")
        if not os.path.exists(project_file):
            return f'Proje "{project_name}" bulunamadı.', 404
        try:
            with open(project_file, 'r', encoding='utf-8') as f:
                project_data = json.load(f)
            stores_list = project_data.get('stores', [])
            if not stores_list:
                return "Projede mağaza verisi bulunamadı.", 404
            base_df = pd.DataFrame(stores_list)
            schedule_data = project_data.get('schedule', {})
            daily_order = project_data.get('daily_order', {})
            # Haritada eklenen güncel notları kullan; formda yoksa dosyadaki store_notes
            notes_data = json.loads(request.form.get('notes_data', '{}')) if request.form.get('notes_data') else project_data.get('store_notes', {})
            edited_stores_data = project_data.get('edited_stores', {})
            new_stores_data = [s for s in stores_list if s.get('isNew')]
            download_type = request.form.get('download_type', 'full_plan')
        except Exception as e:
            return f"Proje dosyası okunurken hata oluştu: {e}", 500
    else:
        try:
            # YENİ: Önce frontend'den gelen stores_data'yı kullanmayı dene
            stores_data_json = request.form.get('stores_data', None)
            
            if stores_data_json:
                # Frontend'den gelen mağaza verileri varsa bunları kullan
                stores_list = json.loads(stores_data_json)
                base_df = pd.DataFrame(stores_list)
            elif LAST_UPLOADED_FILE_PATH and os.path.exists(LAST_UPLOADED_FILE_PATH):
                # Fallback: Eğer stores_data yoksa, eski yöntemi kullan
                base_df = pd.read_csv(LAST_UPLOADED_FILE_PATH, dtype=str, encoding='utf-8')
            else:
                return "İndirilecek veri bulunamadı. Lütfen önce bir Excel dosyası yükleyin veya bir proje yükleyin.", 404
        except Exception as e:
            return f"Veri okunurken hata oluştu: {e}", 500

        download_type = request.form.get('download_type', 'full_plan')
        schedule_data = json.loads(request.form['schedule_data'])
        daily_order = json.loads(request.form.get('daily_order', '{}'))
        notes_data = json.loads(request.form.get('notes_data', '{}'))
        edited_stores_data = json.loads(request.form.get('edited_stores_data', '{}'))
        # YENİ: Yeni eklenen mağazaları al
        new_stores_data = json.loads(request.form.get('new_stores_data', '[]'))

    # --- GÜNCELLEME: Tüm tarihleri topla ve sırala (Gün hesaplaması için) ---
    all_scheduled_dates = set()
    
    # 1. Mevcut planlama tarihlerini topla
    for dates in schedule_data.values():
        for d in dates:
            if d: all_scheduled_dates.add(d)
            
    # 2. Yeni eklenen mağazaların tarihlerini topla
    for store in new_stores_data:
        if store.get('Atanan Tarih'):
            all_scheduled_dates.add(store.get('Atanan Tarih'))

    # 3. Veri tipini algıla ve sırala (Tarih mi yoksa 'X. Gün' formatı mı?)
    def smart_sort_key(val):
        val_str = str(val).strip()
        # Tarih formatı kontrolü (GG.AA.YYYY)
        try:
            parts = val_str.split('.')
            if len(parts) == 3:
                return (0, datetime(int(parts[2]), int(parts[1]), int(parts[0])))
        except:
            pass
        
        # Gün formatı kontrolü (1. Gün, 2, 5. Gün vb.)
        try:
            # Sadece sayıları çek
            num = int(''.join(filter(str.isdigit, val_str)))
            return (1, num)
        except:
            # Bilinmeyen format, sona at
            return (2, val_str)

    sorted_unique_dates = sorted(list(all_scheduled_dates), key=smart_sort_key)
    
    # 4. Haritalama mantığı
    date_day_map = {}
    is_date_based = False
    
    # Verilerin tarih mi gün mü olduğuna karar ver (İlk geçerli veriye bak)
    for d in sorted_unique_dates:
        if isinstance(smart_sort_key(d)[1], datetime):
            is_date_based = True
            break
            
    if is_date_based:
        # Eğer sistem tarih bazlı çalışıyorsa, tarihleri 1. Gün, 2. Gün diye numaralandır
        date_day_map = {date: f"{i+1}. Gün" for i, date in enumerate(sorted_unique_dates)}
    else:
        # Eğer sistem zaten "Gün" bazlı çalışıyorsa (örn: "3. Gün"), olduğu gibi bırak
        # Haritalamaya gerek yok, kendi değeri geçerli olacak
        pass

    # 1. Adım: Ana veri tabanını tarayıcıdaki değişikliklerle güncelle
    if edited_stores_data:
        for store_name, edits in edited_stores_data.items():
            if store_name in base_df['Mağaza İsmi'].values:
                for col, value in edits.items():
                    if col in base_df.columns:
                        base_df.loc[base_df['Mağaza İsmi'] == store_name, col] = str(value)

    # 2. Adım: Yeni planlama sütunlarını eklemeden önce eski planlama sütunlarını temizle
    plan_related_cols = ['Atanan Tarih', 'Sıra', 'Önceki Mağazaya Mesafe (km)', 'Önceki Mağazaya Süre (dk)', 'Açıklama']
    for col in plan_related_cols:
        if col in base_df.columns:
            base_df = base_df.drop(columns=[col])

    # 3. Adım: Her mağaza-tarih kombinasyonu için ayrı satır oluştur
    output_rows = []
    
    # Planlanmış mağazalar için her tarih için ayrı satır
    for store_name, dates_list in schedule_data.items():
        store_data = base_df[base_df['Mağaza İsmi'] == store_name]
        if not store_data.empty:
            store_row = store_data.iloc[0].copy()
            
            # Mağaza İçin Süre'yi numeric'e çevir
            store_duration = 0
            if 'Mağaza İçin Süre' in store_row and store_row['Mağaza İçin Süre']:
                try:
                    store_duration = float(store_row['Mağaza İçin Süre'])
                except:
                    store_duration = 0
            
            # Her tarih için ayrı satır oluştur
            for date in dates_list:
                new_row = store_row.copy()
                new_row['Atanan Tarih'] = date
                
                # GÜNCELLEME: Gün (Plan Tarihi) bilgisini güncelle
                # Eğer tarih bazlıysa hesaplanan günü yaz, değilse mevcut değeri (Örn: "5. Gün") koru
                if date in date_day_map:
                    new_row['Plan Tarihi'] = date_day_map[date]
                elif not is_date_based:
                     # Gün bazlı sistemde "Plan Tarihi" ile "Atanan Tarih" aynıdır veya kullanıcı ne girdiyse odur
                    new_row['Plan Tarihi'] = date
                
                # Bu tarih için rota bilgilerini al
                merch = new_row.get('Personel', '')
                order_list = daily_order.get(date, {}).get(str(merch), [])
                
                if store_name in order_list:
                    sira = order_list.index(store_name) + 1
                    new_row['Sıra'] = sira
                    
                    # Önceki mağazaya mesafe ve süreyi hesapla
                    if sira > 1:
                        prev_store_name = order_list[sira - 2]  # 0-indexed olduğu için -2
                        prev_store_data = base_df[base_df['Mağaza İsmi'] == prev_store_name]
                        if not prev_store_data.empty:
                            prev_store = prev_store_data.iloc[0]
                            try:
                                mesafe, sure = get_route_details(
                                    float(prev_store['Enlem']), float(prev_store['Boylam']),
                                    float(new_row['Enlem']), float(new_row['Boylam'])
                                )
                                new_row['Önceki Mağazaya Mesafe (km)'] = mesafe
                                new_row['Önceki Mağazaya Süre (dk)'] = sure
                            except:
                                new_row['Önceki Mağazaya Mesafe (km)'] = 0.0
                                new_row['Önceki Mağazaya Süre (dk)'] = 0.0
                        else:
                            new_row['Önceki Mağazaya Mesafe (km)'] = 0.0
                            new_row['Önceki Mağazaya Süre (dk)'] = 0.0
                    else:
                        new_row['Önceki Mağaza Mesafe (km)'] = 0.0
                        new_row['Önceki Mağaza Süre (dk)'] = 0.0
                else:
                    new_row['Sıra'] = ''
                    new_row['Önceki Mağazaya Mesafe (km)'] = 0.0
                    new_row['Önceki Mağazaya Süre (dk)'] = 0.0
                
                # Mağaza İçin Süre'yi ekle
                new_row['Mağaza İçin Süre'] = store_duration
                
                # Notları ekle
                new_row['Açıklama'] = notes_data.get(store_name, '')
                
                output_rows.append(new_row)
    
    # YENİ: Atanamayan mağazalar için ayrı liste oluştur
    unassigned_rows = []
    planned_stores = set(schedule_data.keys())
    all_stores = set(base_df['Mağaza İsmi'].tolist())
    unplanned_stores = all_stores - planned_stores
    
    for store_name in unplanned_stores:
        store_data = base_df[base_df['Mağaza İsmi'] == store_name]
        if not store_data.empty:
            store_row = store_data.iloc[0].copy()
            store_row['Atanan Tarih'] = ''
            store_row['Sıra'] = ''
            store_row['Önceki Mağazaya Mesafe (km)'] = ''
            store_row['Önceki Mağazaya Süre (dk)'] = ''
            
            # Mağaza İçin Süre'yi numeric'e çevir
            if 'Mağaza İçin Süre' in store_row and store_row['Mağaza İçin Süre']:
                try:
                    store_row['Mağaza İçin Süre'] = float(store_row['Mağaza İçin Süre'])
                except:
                    store_row['Mağaza İçin Süre'] = 0
            
            store_row['Açıklama'] = notes_data.get(store_name, '')
            unassigned_rows.append(store_row)

    # 4. Adım: DataFrame oluştur
    output_df = pd.DataFrame(output_rows)
    unassigned_df = pd.DataFrame(unassigned_rows)

    # 5. Adım: İndirme tipine göre filtreleme
    if download_type == 'revisions_only':
        # Sadece planlama yapılan mağazaları al (tarihi olanlar)
        output_df = output_df[output_df['Atanan Tarih'] != ''].copy()
        filename = 'revize_edilmis_plan.xlsx'
        sheet_name = 'Revize Edilenler'
    else:
        filename = 'detayli_planlama.xlsx'
        sheet_name = 'Tam Plan'

    # 6. Adım: Excel'i oluştur ve gönder
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Ana plan sayfası
        output_df.to_excel(writer, index=False, sheet_name=sheet_name)
        
        # YENİ: Atanamayanlar sayfası (sadece atanamayan mağazalar varsa)
        if not unassigned_df.empty:
            unassigned_df.to_excel(writer, index=False, sheet_name='Atanamayanlar')
        
        # YENİ: Yeni eklenen mağazalar sayfası
        if new_stores_data:
            new_stores_rows = []
            
            for new_store in new_stores_data:
                store_name = new_store.get('Mağaza İsmi', '')
                if not store_name:
                    continue
                    
                # Bu mağazanın atanmış tarihlerini bul
                assigned_dates = schedule_data.get(store_name, [])
                
                if assigned_dates:
                    # Her tarih için ayrı satır oluştur
                    for date in assigned_dates:
                        new_row = {
                            'Mağaza İsmi': store_name,
                            'Enlem': new_store.get('Enlem', ''),
                            'Boylam': new_store.get('Boylam', ''),
                            'Personel': new_store.get('Personel', ''),
                            'Bölge': new_store.get('Bölge', ''),
                            'Atanan Tarih': date,
                            'Plan Tarihi': date_day_map.get(date, new_store.get('Plan Tarihi', '')),
                            'Mağaza İçin Süre': new_store.get('Mağaza İçin Süre', ''),
                            'İl': new_store.get('İl', ''),
                            'İlçe': new_store.get('İlçe', ''),
                            'Frekans': new_store.get('Frekans', ''),
                            'Açıklama': notes_data.get(store_name, '')
                        }
                        new_stores_rows.append(new_row)
                else:
                    # Tarih atanmamışsa tek satır oluştur
                    new_row = {
                        'Mağaza İsmi': store_name,
                        'Enlem': new_store.get('Enlem', ''),
                        'Boylam': new_store.get('Boylam', ''),
                        'Personel': new_store.get('Personel', ''),
                        'Bölge': new_store.get('Bölge', ''),
                        'Atanan Tarih': '',
                        'Mağaza İçin Süre': new_store.get('Mağaza İçin Süre', ''),
                        'İl': new_store.get('İl', ''),
                        'İlçe': new_store.get('İlçe', ''),
                        'Frekans': new_store.get('Frekans', ''),
                        'Açıklama': notes_data.get(store_name, '')
                    }
                    new_stores_rows.append(new_row)
            
            if new_stores_rows:
                new_stores_df = pd.DataFrame(new_stores_rows)
                new_stores_df.to_excel(writer, index=False, sheet_name='Yeni eklenenler')
        
        # İstatistik sayfası (sadece tam plan için ve planlama varsa)
        if download_type == 'full_plan' and schedule_data:
            # Mağaza bazlı ziyaret sayıları
            visit_counts = {}
            for store_name, dates in schedule_data.items():
                visit_counts[store_name] = len(dates)
            
            if visit_counts:
                visit_df = pd.DataFrame(list(visit_counts.items()), columns=['Mağaza İsmi', 'Toplam Ziyaret Sayısı'])
                visit_df = visit_df.sort_values('Toplam Ziyaret Sayısı', ascending=False)
                visit_df.to_excel(writer, index=False, sheet_name='Mağaza Ziyaret Sayıları')
                
                # Tarih bazlı özet
                date_summary = {}
                for store_name, dates in schedule_data.items():
                    for date in dates:
                        if date not in date_summary:
                            date_summary[date] = 0
                        date_summary[date] += 1
                
                if date_summary:
                    date_df = pd.DataFrame(list(date_summary.items()), columns=['Tarih', 'Toplam Ziyaret Sayısı'])
                    date_df = date_df.sort_values('Tarih')
                    date_df.to_excel(writer, index=False, sheet_name='Tarih Bazlı Özet')

    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=filename)

@app.route('/download_template')
def download_template():
    # Excel kalıbını oluştur
    output = io.BytesIO()
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Kalıp"

    # Başlıkları yaz
    headers = ['Mağaza İsmi', 'Enlem', 'Boylam', 'Personel', 'Mağaza İçin Süre', 'Plan Tarihi', 'İl', 'Bölge', 'Frekans', 'İlçe']
    sheet.append(headers)

    # Örnek veri ekle
    example_data = ['Örnek Mağaza', '41.0082', '28.9784', 'Personel 1', '30', '01.01.2024', 'İstanbul', 'Marmara', 'Haftalık', 'Kadıköy']
    sheet.append(example_data)

    # Dosyayı kaydet
    workbook.save(output)
    output.seek(0)

    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                     as_attachment=True, download_name='kalip.xlsx')

# YENİ: ROUTING SİSTEMİ ROUTE'LARI
@app.route('/routing')
def routing_index():
    return render_template('routing.html')

@app.route('/routing/download-template')
def routing_download_template():
    """Excel şablonunu indir"""
    # Şablon verisi oluştur
    template_data = {
        'Mağaza İsmi': [
            'Migros Bakırköy',
            'Carrefour Avcılar', 
            'Watsons Bahçelievler',
            'Migros Florya',
            'Carrefour Ataköy'
        ],
        'Mağaza İçin Süre': [120, 90, 60, 180, 150],
        'Enlem': [40.99144, 40.9953, 40.99748, 41.01411, 41.0254],
        'Boylam': [28.88108, 28.9097, 28.88641, 28.94547, 29.0432]
    }
    
    template_df = pd.DataFrame(template_data)
    
    # Açıklama sayfası ekle
    explanation_data = {
        'Sütun Adı': [
            'Mağaza İsmi',
            'Mağaza İçin Süre', 
            'Enlem',
            'Boylam'
        ],
        'Açıklama': [
            'Mağazanın adı (Migros, Carrefour, Watsons vb.)',
            'Mağaza ziyaret süresi (dakika cinsinden)',
            'Mağazanın enlem koordinatı (ondalık derece)',
            'Mağazanın boylam koordinatı (ondalık derece)'
        ],
        'Örnek Değer': [
            'Migros Bakırköy',
            '120',
            '40.99144', 
            '28.88108'
        ],
        'Notlar': [
            'Zorunlu alan',
            'Zorunlu alan, sayısal değer',
            'Zorunlu alan, ondalık sayı',
            'Zorunlu alan, ondalık sayı'
        ]
    }
    
    explanation_df = pd.DataFrame(explanation_data)
    
    # Excel dosyasını hafızada oluştur
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        template_df.to_excel(writer, sheet_name='Örnek_Veri', index=False)
        explanation_df.to_excel(writer, sheet_name='Açıklamalar', index=False)
    
    output.seek(0)
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='rota_optimizasyon_sablon.xlsx'
    )

@app.route('/routing/upload', methods=['POST'])
def routing_upload_file():
    if 'file' not in request.files:
        return "Dosya seçilmedi"
    
    file = request.files['file']
    personel_sayisi = int(request.form['personel_sayisi'])
    osrm_url = request.form.get('osrm_url', 'http://localhost:5000/route/v1/driving/')
    
    if file.filename == '':
        return "Dosya seçilmedi"
    
    try:
        print("Excel dosyası okunuyor...")
        stores_df = pd.read_excel(file)
        print(f"Excel okundu. {len(stores_df)} satır bulundu.")
        
        required_columns = ['Mağaza İsmi', 'Mağaza İçin Süre', 'Enlem', 'Boylam']
        missing_columns = [col for col in required_columns if col not in stores_df.columns]
        if missing_columns:
            return f"Excel dosyasında şu sütunlar eksik: {', '.join(missing_columns)}"
        
        stores_df = stores_df.dropna(subset=required_columns)
        print(f"Temizlenmiş veri: {len(stores_df)} satır")
        
        sistem = RotaOptimizasyonSistemi(osrm_url)
        optimized_df = sistem.optimize_routes(stores_df, personel_sayisi)
        
        print(f"Optimizasyon tamamlandı. Sonuç: {len(optimized_df)} satır")
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            optimized_df.to_excel(writer, index=False, sheet_name='Optimize Rotalar')
        
        output.seek(0)
        
        print("Excel çıktısı hazırlandı.")
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'optimize_rotalar_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )
        
    except Exception as e:
        import traceback
        error_details = traceback.format_exc()
        print(f"Hata: {str(e)}")
        return f"Hata oluştu: {str(e)}\n\nDetaylar:\n{error_details}"

@app.route('/api/auto_optimize', methods=['POST'])
def api_auto_optimize():
    try:
        data = request.get_json()
        stores_data = data.get('stores', [])
        start_date = data.get('start_date')
        start_day = data.get('start_day')
        cycle_days = data.get('cycle_days', 28)
        
        if not stores_data:
            return jsonify({'status': 'error', 'message': 'İşlenecek mağaza verisi yok.'})

        # JSON'dan DataFrame'e
        df = pd.DataFrame(stores_data)
        
        # Sayısal çevrimler
        df['Enlem'] = pd.to_numeric(df['Enlem'], errors='coerce')
        df['Boylam'] = pd.to_numeric(df['Boylam'], errors='coerce')
        # GÜNCELLEME: Mağaza İçin Süre sütununu da numeric'e çevir
        if 'Mağaza İçin Süre' in df.columns:
            df['Mağaza İçin Süre'] = pd.to_numeric(df['Mağaza İçin Süre'], errors='coerce')
            df['Mağaza İçin Süre'] = df['Mağaza İçin Süre'].fillna(30)  # Varsayılan 30 dakika
        else:
            df['Mağaza İçin Süre'] = 30  # Sütun yoksa varsayılan değer
        df = df.dropna(subset=['Enlem', 'Boylam'])
        
        # Planlayıcıyı çalıştır
        planner = LojistikPlanlayici()
        # Gün bazlı planlama (start_day geldiyse) veya tarih bazlı (start_date geldiyse)
        result_df = planner.process(df, start_date_str=start_date, start_day=start_day, cycle_days=int(cycle_days) if cycle_days else 28)
        
        # Sonucu döndür
        return jsonify({
            'status': 'success',
            'optimized_stores': result_df.fillna('').to_dict('records')
        })
        
    except Exception as e:
        print(f"Hata: {e}")
        return jsonify({'status': 'error', 'message': str(e)})

@app.route('/api/check_connection', methods=['POST'])
def check_connection():
    global OSRM_API_BASE_URL
    data = request.get_json()
    port = data.get('port', '')
    
    def _normalize_osrm_base(u: str) -> str:
        u = str(u or '').strip()
        if not u:
            return ''
        if not (u.startswith('http://') or u.startswith('https://')):
            u = 'https://' + u
        if 'route/v1/driving' not in u:
            if not u.endswith('/'):
                u += '/'
            u += 'route/v1/driving/'
        else:
            if not u.endswith('/'):
                u += '/'
        return u
    
    # Eğer port girilmediyse varsayılan 5000'i dene, girildiyse o portu ayarla
    target_url = ""
    if not port:
        target_url = "http://localhost:5000/route/v1/driving/"
    else:
        # Eğer tam url girildiyse onu al, sadece port girildiyse localhost ile birleştir
        if port.startswith("http"):
            target_url = _normalize_osrm_base(port)
        else:
            target_url = f"http://localhost:{port}/route/v1/driving/"

    # Test isteği gönder (İstanbul merkezli rastgele kısa bir rota)
    test_coords = "28.9784,41.0082;28.9800,41.0100"
    try:
        response = requests.get(f"{target_url}{test_coords}", timeout=5)
        if response.status_code == 200:
            OSRM_API_BASE_URL = target_url
            return jsonify({'status': 'success', 'url': target_url, 'message': 'Bağlantı Başarılı'})
        else:
            return jsonify({'status': 'error', 'message': f'Sunucu hata kodu döndürdü: {response.status_code}'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': f'Bağlantı başarısız: {str(e)}'})

# YENİ: PERSONEL TAKİP SİSTEMİ ROUTE'LARI
@app.route('/personel_takip')
def personel_takip_index():
    return render_template('personel_takip.html')

@app.route('/personel_takip/upload_contact_report', methods=['POST'])
def upload_contact_report():
    """Kontak Bazlı Rapor dosyasını yükle ve işle"""
    try:
        if 'file' not in request.files:
            return jsonify({'status': 'error', 'message': 'Dosya seçilmedi'})
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'status': 'error', 'message': 'Dosya seçilmedi'})
        
        if file and (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
            # Excel'i oku
            df = pd.read_excel(file, dtype=str)
            
            # Sütun isimlerini normalize et (büyük/küçük harf duyarsız)
            df.columns = df.columns.str.strip()
            
            # Gerekli sütunları kontrol et (farklı yazım şekillerini dene)
            required_columns = {
                'sürücü': ['Sürücü', 'sürücü', 'SURUCU', 'Sürücü Adı', 'SürücüAdı'],
                'başlangıç_enlem': ['Başlangıç Enlem', 'başlangıç enlem', 'BASLANGIC_ENLEM', 'Başlangıç_Enlem', 'BaşlangıçEnlem'],
                'başlangıç_boylam': ['Başlangıç Boylam', 'başlangıç boylam', 'BASLANGIC_BOYLAM', 'Başlangıç_Boylam', 'BaşlangıçBoylam'],
                'bitiş_enlem': ['Bitiş Enlem', 'bitiş enlem', 'BITIS_ENLEM', 'Bitiş_Enlem', 'BitişEnlem'],
                'bitiş_boylam': ['Bitiş Boylam', 'bitiş boylam', 'BITIS_BOYLAM', 'Bitiş_Boylam', 'BitişBoylam']
            }
            
            # Sütun eşleştirmesi
            column_mapping = {}
            for key, possible_names in required_columns.items():
                found = False
                for name in possible_names:
                    if name in df.columns:
                        column_mapping[key] = name
                        found = True
                        break
                if not found:
                    return jsonify({
                        'status': 'error',
                        'message': f"'{possible_names[0]}' sütunu bulunamadı. Mevcut sütunlar: {', '.join(df.columns.tolist())}"
                    })
            
            # Verileri temizle ve dönüştür
            result_data = []
            for _, row in df.iterrows():
                try:
                    record = {
                        'Sürücü': str(row[column_mapping['sürücü']]).strip() if pd.notna(row[column_mapping['sürücü']]) else '',
                        'Başlangıç Enlem': str(row[column_mapping['başlangıç_enlem']]).strip() if pd.notna(row[column_mapping['başlangıç_enlem']]) else '',
                        'Başlangıç Boylam': str(row[column_mapping['başlangıç_boylam']]).strip() if pd.notna(row[column_mapping['başlangıç_boylam']]) else '',
                        'Bitiş Enlem': str(row[column_mapping['bitiş_enlem']]).strip() if pd.notna(row[column_mapping['bitiş_enlem']]) else '',
                        'Bitiş Boylam': str(row[column_mapping['bitiş_boylam']]).strip() if pd.notna(row[column_mapping['bitiş_boylam']]) else ''
                    }
                    
                    # Geçersiz kayıtları filtrele
                    if record['Sürücü'] and record['Başlangıç Enlem'] and record['Başlangıç Boylam'] and record['Bitiş Enlem'] and record['Bitiş Boylam']:
                        result_data.append(record)
                except Exception as e:
                    print(f"Satır işleme hatası: {e}")
                    continue
            
            return jsonify({
                'status': 'success',
                'data': result_data,
                'count': len(result_data)
            })
        else:
            return jsonify({'status': 'error', 'message': 'Geçersiz dosya formatı'})
            
    except Exception as e:
        print(f"Kontak rapor yükleme hatası: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'status': 'error', 'message': f'Dosya işlenirken hata oluştu: {str(e)}'})

@app.route('/personel_takip/upload_store_locations', methods=['POST'])
def upload_store_locations():
    """Mağazaların Konumu dosyasını yükle ve işle"""
    try:
        if 'file' not in request.files:
            return jsonify({'status': 'error', 'message': 'Dosya seçilmedi'})
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'status': 'error', 'message': 'Dosya seçilmedi'})
        
        if file and (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
            # Excel'i oku
            df = pd.read_excel(file, dtype=str)
            
            # Sütun isimlerini normalize et
            df.columns = df.columns.str.strip()
            
            # Gerekli sütunları kontrol et (farklı yazım şekillerini dene)
            required_columns = {
                'personel': ['Personel', 'personel', 'PERSONEL', 'Personel Adı', 'PersonelAdı'],
                'enlem': ['Enlem', 'enlem', 'ENLEM', 'Latitude', 'LAT'],
                'boylam': ['Boylam', 'boylam', 'BOYLAM', 'Longitude', 'LON', 'Long']
            }
            
            # Sütun eşleştirmesi
            column_mapping = {}
            for key, possible_names in required_columns.items():
                found = False
                for name in possible_names:
                    if name in df.columns:
                        column_mapping[key] = name
                        found = True
                        break
                if not found:
                    return jsonify({
                        'status': 'error',
                        'message': f"'{possible_names[0]}' sütunu bulunamadı. Mevcut sütunlar: {', '.join(df.columns.tolist())}"
                    })
            
            # Verileri temizle ve dönüştür
            result_data = []
            for _, row in df.iterrows():
                try:
                    record = {
                        'Personel': str(row[column_mapping['personel']]).strip() if pd.notna(row[column_mapping['personel']]) else '',
                        'Enlem': str(row[column_mapping['enlem']]).strip() if pd.notna(row[column_mapping['enlem']]) else '',
                        'Boylam': str(row[column_mapping['boylam']]).strip() if pd.notna(row[column_mapping['boylam']]) else ''
                    }
                    
                    # Geçersiz kayıtları filtrele
                    if record['Personel'] and record['Enlem'] and record['Boylam']:
                        # Koordinatların geçerli sayı olup olmadığını kontrol et
                        try:
                            float(record['Enlem'])
                            float(record['Boylam'])
                            result_data.append(record)
                        except ValueError:
                            continue
                except Exception as e:
                    print(f"Satır işleme hatası: {e}")
                    continue
            
            return jsonify({
                'status': 'success',
                'data': result_data,
                'count': len(result_data)
            })
        else:
            return jsonify({'status': 'error', 'message': 'Geçersiz dosya formatı'})
            
    except Exception as e:
        print(f"Mağaza konum yükleme hatası: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'status': 'error', 'message': f'Dosya işlenirken hata oluştu: {str(e)}'})

# YENİ: PROJE YÖNETİMİ API ENDPOINT'LERİ

# Proje dosyalarının saklanacağı dizin
PROJECTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'saved_projects')
if not os.path.exists(PROJECTS_DIR):
    os.makedirs(PROJECTS_DIR)

@app.route('/api/save_project', methods=['POST'])
def save_project():
    """Projeyi kaydet"""
    try:
        data = request.get_json()
        project_name = data.get('project_name', '').strip()
        
        if not project_name:
            return jsonify({'status': 'error', 'message': 'Proje adı boş olamaz'})
        
        # Geçersiz karakterleri temizle
        safe_name = "".join(c for c in project_name if c.isalnum() or c in (' ', '-', '_')).strip()
        if not safe_name:
            return jsonify({'status': 'error', 'message': 'Geçersiz proje adı'})
        
        project_data = {
            'name': project_name,
            'created_at': data.get('created_at', time.strftime('%Y-%m-%d %H:%M:%S')),
            'last_modified': time.strftime('%Y-%m-%d %H:%M:%S'),
            'csv_data': data.get('csv_data', ''),
            'schedule': data.get('schedule', {}),
            'daily_order': data.get('daily_order', {}),
            'store_notes': data.get('store_notes', {}),
            'edited_stores': data.get('edited_stores', {}),
            'empty_days': data.get('empty_days', {}),
            'stores': data.get('stores', [])
        }
        
        # Proje dosyasını kaydet
        project_file = os.path.join(PROJECTS_DIR, f"{safe_name}.json")
        with open(project_file, 'w', encoding='utf-8') as f:
            json.dump(project_data, f, ensure_ascii=False, indent=2)
        
        return jsonify({
            'status': 'success',
            'message': f'Proje "{project_name}" başarıyla kaydedildi',
            'project_name': project_name
        })
        
    except Exception as e:
        print(f"Proje kaydetme hatası: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'status': 'error', 'message': f'Proje kaydedilemedi: {str(e)}'})

@app.route('/api/load_project', methods=['POST'])
def load_project():
    """Projeyi yükle"""
    try:
        data = request.get_json()
        project_name = data.get('project_name', '').strip()
        
        if not project_name:
            return jsonify({'status': 'error', 'message': 'Proje adı boş olamaz'})
        
        # Geçersiz karakterleri temizle
        safe_name = "".join(c for c in project_name if c.isalnum() or c in (' ', '-', '_')).strip()
        project_file = os.path.join(PROJECTS_DIR, f"{safe_name}.json")
        
        if not os.path.exists(project_file):
            return jsonify({'status': 'error', 'message': f'Proje "{project_name}" bulunamadı'})
        
        # Proje dosyasını oku
        with open(project_file, 'r', encoding='utf-8') as f:
            project_data = json.load(f)
        
        return jsonify({
            'status': 'success',
            'message': f'Proje "{project_name}" başarıyla yüklendi',
            'project_data': project_data
        })
        
    except Exception as e:
        print(f"Proje yükleme hatası: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'status': 'error', 'message': f'Proje yüklenemedi: {str(e)}'})

@app.route('/api/list_projects', methods=['GET'])
def list_projects():
    """Tüm projeleri listele"""
    try:
        projects = []
        
        if not os.path.exists(PROJECTS_DIR):
            return jsonify({'status': 'success', 'projects': []})
        
        for filename in os.listdir(PROJECTS_DIR):
            if filename.endswith('.json'):
                project_file = os.path.join(PROJECTS_DIR, filename)
                try:
                    with open(project_file, 'r', encoding='utf-8') as f:
                        project_data = json.load(f)
                        projects.append({
                            'name': project_data.get('name', filename[:-5]),
                            'created_at': project_data.get('created_at', 'Bilinmiyor'),
                            'last_modified': project_data.get('last_modified', 'Bilinmiyor')
                        })
                except Exception as e:
                    print(f"Proje okuma hatası ({filename}): {e}")
                    continue
        
        # Tarihe göre sırala (en yeni en üstte)
        projects.sort(key=lambda x: x.get('last_modified', ''), reverse=True)
        
        return jsonify({'status': 'success', 'projects': projects})
        
    except Exception as e:
        print(f"Proje listeleme hatası: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'status': 'error', 'message': f'Projeler listelenemedi: {str(e)}'})

@app.route('/api/delete_project', methods=['POST'])
def delete_project():
    """Projeyi sil"""
    try:
        data = request.get_json()
        project_name = data.get('project_name', '').strip()
        
        if not project_name:
            return jsonify({'status': 'error', 'message': 'Proje adı boş olamaz'})
        
        # Geçersiz karakterleri temizle
        safe_name = "".join(c for c in project_name if c.isalnum() or c in (' ', '-', '_')).strip()
        project_file = os.path.join(PROJECTS_DIR, f"{safe_name}.json")
        
        if not os.path.exists(project_file):
            return jsonify({'status': 'error', 'message': f'Proje "{project_name}" bulunamadı'})
        
        # Proje dosyasını sil
        os.remove(project_file)
        
        return jsonify({
            'status': 'success',
            'message': f'Proje "{project_name}" başarıyla silindi'
        })
        
    except Exception as e:
        print(f"Proje silme hatası: {e}")
        import traceback
        traceback.print_exc()
        return jsonify({'status': 'error', 'message': f'Proje silinemedi: {str(e)}'})

if __name__ == '__main__':
    # EXE içinde: UI kapanınca server'ı otomatik kapat (heartbeat yoksa timeout)
    if bool(getattr(sys, "frozen", False)):
        def heartbeat_watcher(timeout_seconds=8, check_every_seconds=2):
            def _loop():
                while True:
                    time.sleep(check_every_seconds)
                    try:
                        if time.time() - LAST_CLIENT_PING_TS > timeout_seconds:
                            os._exit(0)
                    except Exception:
                        os._exit(0)
            from threading import Thread
            Thread(target=_loop, daemon=True).start()

        heartbeat_watcher()

    # Tarayıcıyı otomatik aç
    Timer(1, open_browser).start()
    # Uygulamayı başlat
    app.run(debug=False, host='0.0.0.0', port=5001)
