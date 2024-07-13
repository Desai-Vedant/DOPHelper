import os
import glob
import logging
import pandas as pd


# Configure the logging settings
logging.basicConfig(filename='doplogs.log' ,level=20, format='%(asctime)s - %(levelname)s - %(message)s')

class UpdateTaskError(Exception):
    pass

class DatabaseError(Exception):
    pass

# Does not take any argument while initialize
class DOPDatabaseAssistant:
    def __init__(self):
        self.create_folder_if_not_exists('Database')
        self.create_folder_if_not_exists('RDRecord')
        self.db_path = './Database/Database.csv'
        self.initialize_database(self.db_path)
        self.records_folder_path = './RDRecord/'
        self.db = pd.read_csv(self.db_path, dtype={
                'ac_no': str,
                'ac_id': int
            })
        
    # Creates folder if folder of given name does not exist
    def create_folder_if_not_exists(self, folder_name):
        if not os.path.exists(folder_name):
            os.makedirs(folder_name)
            
    # Initialize the database if database.csv file does not exist
    def initialize_database(self, db_path):
        # Define the columns and their data types
        columns = {'ac_no': str, 'ac_id': int, 'acc_holder_name': str, 'denomination': str, 'acc_opening_date': str,
                'no_of_installments': int, 'is_active': str, 'aslaas_no': str}

        # Create an empty DataFrame with the specified columns
        db = pd.DataFrame(columns=columns)

        # Check if the database file exists
        if not os.path.exists(db_path):
            # Save the DataFrame to the specified file path
            db.to_csv(db_path, index=False)

    # Checks if there are any new accounts data in the records folder and then updates then in database
    def sync_database_task(self):
        try:
            # Define the data types
            data_types = {
                'ac_no': str
            }
            logging.info("Datatypes Defined Sucessful !")

            df = pd.read_csv(self.find_latest_csv(self.records_folder_path), dtype=data_types)
            db = pd.read_csv(self.db_path, dtype={
                'ac_no': str,
                'ac_id':int
            })

            # Drop Specific columns from DataFrame
            df2 = df.drop(columns=['acc_holder_name','denomination','acc_opening_date'],axis=1)
            logging.info("Reading and Processing of files Sucessful !")

            db['no_of_installments'] = db['no_of_installments'].fillna(-1)
            db['no_of_installments'] = db['no_of_installments'].astype(int)

            # Perform left join and update values in 'Common1' and 'Common2' columns
            db.set_index('ac_no', inplace=True)
            df2.set_index('ac_no', inplace=True)
            db.update(df2)

            
            # Reset the index if needed
            db.reset_index(inplace=True)
            logging.info("Updating Old Data Sucessful !")

            temp_df = db[['ac_id','ac_no','aslaas_no']]
            merged_df = pd.merge(df,temp_df, how='left', on='ac_no')
            new_accs = merged_df[merged_df['ac_id'].isna()]
            last_ac_id = db['ac_id'].max()
            if pd.notna(last_ac_id):
                last_ac_id = int(last_ac_id)
            else:
                last_ac_id = 0
            new_accs.loc[:, 'ac_id'] = new_accs.loc[:, 'ac_id'].fillna(-1)
            new_accs.loc[:, 'ac_id'] = new_accs.loc[:, 'ac_id'].astype(int)
            new_accs.loc[:, 'ac_id'] = list(range(int(last_ac_id)+1, int(last_ac_id) + len(new_accs)+1))
            final_db = pd.concat([db, new_accs])
            final_db.loc[:, 'ac_id'] = final_db.loc[:, 'ac_id'].astype(int)
            
            dtypes = {
                'ac_no' : str,
                'aslaas_no' : 'str'
            }
            
            aslaas_data = pd.read_csv('./temp/aslaas_report.csv', dtype=dtypes)
            logging.info("Loading Aslaas Data Sucessful !")
            
            #database = final_db.merge(aslaas_data, how="left", on="ac_no")
            final_db.set_index('ac_no', inplace=True)
            aslaas_data.set_index('ac_no', inplace=True)
            
            final_db.update(aslaas_data)
            final_db.reset_index(inplace=True)
            
            logging.info("Merging Sucessful!")
            # Update aslaas_no in df1 based on merged_df
            #final_db["aslaas_no"] = final_db["aslaas_no"].fillna(database["aslaas_no_y"])
            logging.info("Updating Aslaas Data Sucessful !")
            
            final_db.to_csv(self.db_path, index=False)
            logging.info("Database Save Sucessful !")

        except Exception as e:
            logging.error(f"Error During Database Update : {e}")
            raise UpdateTaskError(f"Error during Updating Database: {e}")

    # Finds the latest csv file in a directory 
    def find_latest_csv(self,directory):
        # Get list of all CSV files in directory
        files = glob.glob(os.path.join(directory, "*.csv"))
        # Return the file with the latest modification time
        return max(files, key=os.path.getmtime)
    
    # Gets all account numbers which do not have a aslaas number
    def get_ac_nos_without_aslaas(self):
        try:
            db = pd.read_csv(self.db_path, dtype={
                'ac_no': str,
                'ac_id':int
            })

            db = db[db['is_active']==1]

            acc_nos = db[db.aslaas_no.isnull()]['ac_no'].tolist()

            return acc_nos
        except Exception as e:
            logging.error(f"Error During Database Task : {e}")
            raise UpdateTaskError(f"Error during Database Task : {e}")

    # Gets Data of all accounts for seeing all accounts page
    def get_all_accounts(self):
        try:
            db = pd.read_csv(self.db_path, dtype={
                'ac_no': str,
                'ac_id': int
            })

            #Change isActive column info for user Convienence
            all_accounts = db
            all_accounts['is_active'] = all_accounts['is_active'].astype(str)
            all_accounts.loc[(all_accounts.loc[:,'is_active']).astype(str)=='1', 'is_active'] = "Active"
            all_accounts.loc[(all_accounts.loc[:,'is_active']).astype(str)=='0', 'is_active'] = "Closed"

            return all_accounts

        except Exception as e:
            logging.error(f"Error occurred while getting all accounts: {e}")
            raise DatabaseError("Failed to retrieve accounts") from e  # Re-raise with a more specific message

    # Gets account numbers using the ids array
    def get_acc_nos_using_ids(self, acc_ids):
        try:
            acc_nos = self.db[self.db['ac_id'].astype(str).isin(acc_ids)]['ac_no'].astype(str).tolist()
            return acc_nos
        except Exception as e:
            logging.error(f"Error occurred while getting account Numbers: {e}")
            raise DatabaseError("Failed to retrieve accounts") from e  # Re-raise with a more specific message

    # Fetch all account ids of active accounts
    def get_list_of_active_account_ids(self):
        try:
            active_acc_ids = self.db[self.db['is_active']==1]['ac_id'].astype(str).tolist()
            return active_acc_ids
        except Exception as e :
            logging.error(f"Error occurred while getting active account ID's: {e}")
            raise DatabaseError("Failed to retrieve account ID's") from e  # Re-raise with a more specific message

    # Gets details of specific accounts using their ids list
    def get_ac_details_by_ids(self,acc_ids):
        try:
            acc_details = self.db[self.db['ac_id'].isin(acc_ids)]
            acc_details = acc_details.drop(columns=['no_of_installments','is_active','aslaas_no'], axis = 1)
            return acc_details
        except Exception as e:
            logging.error(f"Error occurred while getting account Numbers: {e}")
            raise DatabaseError("Failed to retrieve accounts") from e  # Re-raise with a more specific message

    # Get data of accounts using id list
    def get_data_for_declaration(self, acc_ids):
        try:
            acc_details = self.db[self.db['ac_id'].isin(acc_ids)]
            acc_details = acc_details.drop(columns=['ac_id','no_of_installments','is_active','aslaas_no'], axis = 1)
            return acc_details
        except Exception as e:
            logging.error(f"Error occurred while getting account Numbers: {e}")
            raise DatabaseError("Failed to retrieve accounts") from e  # Re-raise with a more specific message
       
    # Method to add an account to database using account details array 
    def add_account_to_database(self, acc_details):
        try:
            acc_id = len(self.db.index)
            self.db.loc[acc_id] = [str(acc_details[0]),int(acc_id+1),acc_details[1],acc_details[2],acc_details[3],acc_details[4],acc_details[5],acc_details[6]]
            self.db.to_csv(self.db_path, index=False)
        except Exception as e:
            logging.error(f"Error occurred while adding account: {e}")
            raise DatabaseError("Failed to add account") from e  # Re-raise with a more specific message
        return acc_id+1
    
    # Update aslaas numbers in the database for fixed given number of accounts and given aslaas numbers
    def sync_aslaas_numbers(self, acc_ids, aslaas_nos):
        try:
            for i in range(len(acc_ids)):
                account_id = int(acc_ids[i])
                self.db.loc[self.db['ac_id'] == account_id, 'aslaas_no'] = str(aslaas_nos[i])
                    
            self.db.to_csv(self.db_path, index=False)
        except Exception as e:
            logging.error(f"Error Occurred During Updating Database : {e}")
            raise UpdateTaskError("Failed to Update Data") from e

    def get_acc_ids_using_nos(self, acc_nos):
        try:
            acc_ids = self.db[self.db['ac_no'].isin(acc_nos)]['ac_id'].astype(int).tolist()
            return acc_ids
        except Exception as e:
            logging.error(f"Error occurred while getting account IDs: {e}")
            raise DatabaseError("Failed to retrieve account IDs") from e  # Re-raise with a more specific message 
            
    
