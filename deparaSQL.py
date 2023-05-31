# -*- coding: utf-8 -*-
import json
import pandas as pd
from sqlalchemy import create_engine
from extraUtils import gap

def updateSQL():
    try:
        # Load the database connection information from the config.json file
        with open(gap('sec\\config.json'), 'r') as f:
            config = json.load(f)
        
        # Create an engine that connects to the database
        engine = create_engine(f'mysql+mysqlconnector://{config["dbUsername"]}:{config["dbPassword"]}@{config["dbHost"]}:{config["dbPort"]}/{config["dbName"]}')
        
        # Read the data from the Excel file
        df = pd.read_excel('X:\\05 - PROJETOS\\PGD-SHAREPOINT\\De-Para.xlsx', keep_default_na=False)
        
        # Replace cells that contain '-' with empty cells
        df = df.replace('-', '')
        
        # Write the data to the SQL table
        df.to_sql('De-Para', con=engine, if_exists='replace', schema='SISGP', index=False)
    except Exception as e:
        print(f'An error occurred: {e}')

if __name__ == '__main__':
    updateSQL()
