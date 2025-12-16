import sys, argparse, base64, csv, io
import pyodbc  # ensure installed: pip install pyodbc

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--conn", required=True)
    parser.add_argument("--query", required=True)
    parser.add_argument("--filename", required=True)
    args = parser.parse_args()

    # Connect to SQL Server
    conn = pyodbc.connect(args.conn)
    cursor = conn.cursor()
    cursor.execute(args.query)

    # Write results to CSV in memory
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow([column[0] for column in cursor.description])  # headers
    for row in cursor.fetchall():
        writer.writerow(row)

    conn.close()

    # Encode CSV
    encoded = base64.b64encode(output.getvalue().encode("utf-8")).decode("ascii")
    print(f"{args.filename}|{encoded}")

if __name__ == "__main__":
    main()
