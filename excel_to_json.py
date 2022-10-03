import pandas as pd
import sys
import uuid

class excelToJSON:

    def __init__(self) -> None:
        # Check to make sure the command line arguements 
        # are atleast 2 arguements
        if len(sys.argv[1:]) != 2:
            print(f'ERROR: Check your command line arguments!,\n 2 arguements expected [excel_file, output_file]')
            return None
            
        self.in_file = sys.argv[1:][0]
        self.out_file = sys.argv[1:][1]
        df = self.excel_to_pd()
        self.json_data = self.pd_to_json(df)
        self.write_json(self.json_data)

    def excel_to_pd(self):
        # Convert the excel data to pd DataFrame for processing
        df = pd.read_excel(self.in_file)

        '''
        START Data Cleanse: manupilate data to suit your needs. Below is just an example. Not necessary
        '''
        # Rename the columns
        df.rename(columns={"Name":"name", "Email Address": 'email', "Phone": 'phone', "NOTES": 'notes', "TAGS": 'tags'}, inplace=True)
        
        # Generates an unique ID for each set of data using its email and timestamp.
        # Extracts the hex value.
        # Creates a new key 'uid'.
        # Stores the unique ID as the value for uid.
        for num in df.index:
            df.loc[num, 'uid'] = uuid.uuid3(uuid.NAMESPACE_DNS, df.loc[num,'email']).hex
        
        '''
        END Example
        '''

        # Returns the processed df
        return df


    def pd_to_json(self, df):
        '''
        Converts the finalized df to json
        '''
        return df.to_json(orient='records')
    
    def write_json(self, data):
        '''
        Outputs the data in json format
        '''
        with open(self.out_file, 'w') as output:
            output.write(data)


if __name__ == "__main__":
    excelToJSON()