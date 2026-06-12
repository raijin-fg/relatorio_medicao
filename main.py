import sys
from pathlib import Path
from dotenv import load_dotenv
from extract_zabbix import extract_zabbix_data
from extract_netapp import extract_netapp_data
from extract_veritas import extract_netbackup_data


def main():
    """Populate database with current data from all sources."""
    env_path = Path(__file__).parent / ".env"
    if not env_path.exists():
        print(f"Error: .env file not found at {env_path}")
        print("Please create a .env file with the required credentials.")
        return 1

    load_dotenv(env_path)

    print("=" * 60)
    print("DATABASE POPULATION SCRIPT")
    print("=" * 60)

    results = {}

    print("\n[1/3] Extracting Zabbix availability data...")
    results["zabbix"] = extract_zabbix_data()

    print("\n[2/3] Extracting NetApp volume data...")
    results["netapp"] = extract_netapp_data()

    print("\n[3/3] Extracting NetBackup data...")
    results["netbackup"] = extract_netbackup_data()

    print("\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)
    for source, success in results.items():
        status = "✓ SUCCESS" if success else "✗ FAILED"
        print(f"{source.upper():.<15} {status}")

    all_success = all(results.values())
    print("=" * 60)
    if all_success:
        print("All data sources populated successfully!")
        return 0
    else:
        print("Some data sources failed. Check the output above for details.")
        return 1


if __name__ == "__main__":
    sys.exit(main())
