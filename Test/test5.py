import os
import time
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderUnavailable, GeocoderServiceError

def get_coordinates(address, max_retries=3, retry_delay=5):
    retries = 0
    while retries < max_retries:
        try:
            # 设置正确的代理（根据实际情况修改下面的代理地址和端口）
            os.environ['http_proxy'] = "http://proxy.example.com:8080"
            os.environ['https_proxy'] = "http://proxy.example.com:8080"

            geolocator = Nominatim(user_agent="my_app")
            location = geolocator.geocode(address)
            if location:
                return location.latitude, location.longitude
            return None, None
        except (GeocoderUnavailable, GeocoderServiceError) as e:
            print(f"获取坐标时出错: {e}，正在重试...")
            retries += 1
            time.sleep(retry_delay)
    print("多次重试后仍无法获取地理编码服务，请稍后再试。")
    return None, None

# 测试示例
address = "北京市天安门广场"
latitude, longitude = get_coordinates(address)
if latitude and longitude:
    print(f"地址 {address} 的坐标为：纬度 {latitude}，经度 {longitude}")
else:
    print(f"无法获取地址 {address} 的坐标信息")