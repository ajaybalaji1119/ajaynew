import requests
import json
import pandas as pd

API_KEY = "VIRUSTOTAL_API_KEY"

def get_sample_info(hash):
    url = f"https://www.virustotal.com/api/v3/files/{hash}"
    headers = {
        "x-apikey": API_KEY
    }
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json()
    else:
        return None

def save_results_to_excel(results, input_file):
    df = pd.DataFrame(results)
    output_file = input_file.split(".")[0] + "_results.xlsx"
    df.to_excel(output_file, index=False)
    print(f"Results saved to {output_file}")

def main():
    input_file = "sample_hashes.xlsx"
    df = pd.read_excel(input_file)
    sample_hashes = df["Sample Hash"].tolist()
    results = []

    for sample_hash in sample_hashes:
        sample_info = get_sample_info(sample_hash)
        if sample_info:
            result = {
                "Sample Hash": sample_hash,
                "First Seen": sample_info['data']['attributes']['first_submission_date'],
                "Last Seen": sample_info['data']['attributes']['last_analysis_date'],
                "Detections": sample_info['data']['attributes']['last_analysis_stats']['malicious']
            }
            results.append(result)
        else:
            print(f"No information found for sample hash: {sample_hash}")

    save_results_to_excel(results, input_file)

if __name__ == '__main__':
    main()
