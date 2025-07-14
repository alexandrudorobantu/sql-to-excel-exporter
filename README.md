# SQL to Excel Exporter

This project provides a simple way to export order data from a SQL database to an Excel file. It utilizes Python libraries such as `pandas` and `SQLAlchemy` to handle data retrieval and exporting.

## Project Structure

```
sql-to-excel-exporter
├── src
│   ├── export.py                # Main script to execute SQL and export to Excel
│   └── sql
│       └── order_data_query.sql # SQL script to retrieve order data
├── requirements.txt             # Project dependencies
└── README.md                    # Project documentation
```

## Setup Instructions

1. **Clone the repository:**
   ```bash
   git clone <repository-url>
   cd sql-to-excel-exporter
   ```

2. **Create a virtual environment (optional but recommended):**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows use `venv\Scripts\activate`
   ```

3. **Install the required packages:**
   ```bash
   pip install -r requirements.txt
   ```

## Configuration

Before running the export script, ensure that the SQL connection parameters in `export.py` are correctly set up to connect to your database.

## Running the Export Script

To execute the export process, run the following command:

```bash
python src/export.py
```

This will connect to the SQL database, execute the query defined in `order_data_query.sql`, and export the results to an Excel file.

## Dependencies

The project requires the following Python packages:

- `pandas`
- `SQLAlchemy`
- Database driver for your specific SQL database (e.g., `pyodbc` for SQL Server)

Make sure to include any additional dependencies in the `requirements.txt` file as needed.

## License

This project is licensed under the MIT License - see the LICENSE file for details.