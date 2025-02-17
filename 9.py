import requests
import json
from datetime import datetime

# Configuration
PROMETHEUS_URL = "http://localhost:9090"  # Prometheus URL
QUERY_RANGE_ENDPOINT = f"{PROMETHEUS_URL}/api/v1/query_range"

PANEL_QUERIES = {
    "Quick CPU / Mem / Disk": [
        "100 - (avg(irate(node_cpu_seconds_total{mode=\"idle\"}[5m])) * 100)"  # CPU usage in percentage
    ]
}

# Time range (last 5 minutes)
now = int(datetime.now().timestamp())
start_time = now -1  # 5 minutes ago
end_time = now

# Function to query Prometheus metrics
def query_metrics(query):
    params = {
        "query": query,
        "start": start_time,
        "end": end_time,
        "step": 30  # Query step interval in seconds
    }
    response = requests.get(QUERY_RANGE_ENDPOINT, params=params)
    
    if response.status_code != 200:
        print(f"Error querying {query}: {response.status_code} - {response.text}")
        return {}
    
    return response.json()

# Function to convert timestamps to human-readable format
def convert_timestamp(timestamp):
    return datetime.utcfromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M:%S')

# Function to fetch CPU usage percentage and format it
def fetch_cpu_usage():
    query = PANEL_QUERIES["Quick CPU / Mem / Disk"][0]  # CPU usage query
    result = query_metrics(query)
    formatted_data = []

    if "data" in result and "result" in result["data"]:
        for res in result["data"]["result"]:
            metric_values = res["values"]
            
            for ts, value in metric_values:
                try:
                    # Convert value to float and format it as a percentage
                    percentage = f"{float(value):.2f} %"
                    formatted_data.append([convert_timestamp(float(ts)), percentage])
                except ValueError:
                    print(f"Error converting value: {value}")

    return formatted_data

# Save the data with proper structure into JSON format
def save_data(cpu_data):
    data = {
        "Quick CPU / Mem / Disk": {
            "Cpu Busy": cpu_data
        }
    }
    with open("cpu_usage_data_promql.json", "w") as file:
        json.dump(data, file, indent=4)

# Run the script
cpu_data = fetch_cpu_usage()
save_data(cpu_data)
print("CPU usage data extracted and saved to cpu_usage_data_promql.json!")