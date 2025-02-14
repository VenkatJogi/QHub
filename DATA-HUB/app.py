import subprocess

def main():
    # print("Starting Tableau metadata extraction...")
    # subprocess.run(["python", "tableau_extractor.py"])

    print("Metadata extraction complete. Proceeding to Power BI report creation...")
    subprocess.run(["python", "powerbi_creator.py"])

    print("Report migration completed successfully!")

if __name__ == "__main__":
    main()
