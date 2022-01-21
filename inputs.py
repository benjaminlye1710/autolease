### Import from dependencies
from dependencies import *

# Initialise the excel misc functions to use
Excel_Misc_Fns = ExcelHelper()
XL_Fns = Excel_Misc_Fns

### Classes

# =============================================================================
#### Input class
# =============================================================================
class LeaseDataReader(ExcelHelper):
    '''
    Reads in the data from the standardised template.
    
    This template will be prepared by the engagement team.
    '''
    
    HEADER_STR_COL = [
        'Country',
        'Company',
        'Contract',
        'New Contract',
        'Renew Contract?',
        'Is there more than one remeasurement throughout the contract period life?',
        'Early Termination? (date)',
        ]
    
    HEADER_NUM_COL = [
        'S/N',
        'Opening ER (PFY)',
        'Average ER (PFY)',
        'Opening ER',
        'Average ER',
        'Closing ER',
        'Reinstatement Cost (PFY)',
        'Rental Deposit',
        'Opening ROU (PFY)',
        'Opening Accumulated Depreciation (PFY)',
        'Opening Lease Liability (PFY)',
        'Remeasurements',
        'Others (PFY)',
        'Reinstatement Cost',
        'Opening ROU',
        'Opening Accumulated Depreciation',
        'Opening Lease Liability',
        'Others',
        ]

    LINE_COL = [
        'Borrowing Rate (PFY)',
        'Rental/mth (PFY)',
        'Lease Start (PFY)',
        'Lease End (PFY)',
        'Period (months) (PFY)',
        'Option to renew ?',
        'Borrowing Rate',
        'Rental/mth',
        'Lease Start',
        'Lease End',
        'Period (months)',
        'Remarks'
        ]
    
    EXPECTED_COL = HEADER_STR_COL + HEADER_NUM_COL + LINE_COL

    
    def __init__(self, input_fp, sheet_name = 'Lease Data',):
        """
        Generate a dictionary containing:
            - Interest Rate: effective interest rate
            - Client: name of client
            - Current FY End: client's FY End date for current year
            - FX to SGD Conversion Rates
            - df: Pandas DataFrame containing the following data for each lease:
                    - Company    
                    - Branch Code
                    - Country
                    - Prev FY Opening ER
                    - Prev FY Average ER
                    - Prev FY Closing ER
                    - Curr FY Average ER
                    - Curr FY Closing ER
                    - Old Reinstatement Cost
                    - Old Lease Start: date of last day of lease
                    - Old Lease End: date of last day of lease
                    - Old Rental/mth
                    - Old borrowing rate
                    - New contract?
                    - Renew contract? 
                    - Early terminaton? (date)
                    - Option to renew?
                    - Reinstatement cost 
                    - Lease Start: date of last day of lease
                    - Lease End: date of last day of lease
                    - Rental/mth
                    - Borrowing rate                    
        """
        
        # save attributes
        self.input_fp = input_fp
        self.sheet_name = sheet_name
                
        # run methods  
    
    def __main__(self):
        
        self.get_raw_df()
        self.get_meta_info()
        self.get_main_data()
        self.process_data()
        # self.get_all_contracts()
    
    def get_raw_df(self):
        
        # Reads the data
        df0 = pd.read_excel(self.input_fp, self.sheet_name, header = None)
        
        # Converts to excel rows and columns notations
        df0.index = range(1, df0.shape[0]+1)
        df0.columns = (
            (df0.columns + 1)
            .map(openpyxl.utils.get_column_letter)
            )
        
        # Drop empty rows
        df0 = df0.dropna(how='all')
        
        # Save attribute
        self.df0  = df0.copy()
    
    def get_meta_info(self):
        
        df0 = self.df0.copy()
        
        # Get the client name        
        client = df0.at[1, 'A']
        
        # Get the fy end date from the following example string:
        #     - 'Audit for the financial year ended 31 December 2019'
        fy_end = df0.at[2, 'A']
        regex_pat = '\d{1,2} \w+ \d{2,4}'

        try:
            fy_end = re.search(regex_pat, fy_end).group()
        except AttributeError:
            msg = (
                f'Unexpected date format {fy_end} in cell A2. '
                'Please use the format: DD MMM YYYY.'
                )
            raise Exception(msg)
        
        # Check and convert to timestamp
        ensure_correct_date_format(fy_end,'dmy')
        
        fy_end = (
            pd.to_datetime(fy_end, dayfirst = True)
            )
        
        # retrieve fy start date
        fy_start = (shift_month(pd.to_datetime(fy_end), 
                                -11, month_begin = True)
                    )
        
        # retrieve prev fy start date 
        pfy_start = (shift_month(pd.to_datetime(fy_end),
                                     -23, month_begin = True)
                         )
        
        # retrieve prev fy end date
        pfy_end = (shift_month(pd.to_datetime(fy_end),
                                     -12)
                         )
        
        # retrieve schedule starting period
        # pfy start date = 2019, 7, 1 -> 2019, 1, 1
        cutoff_year = (
            shift_month(pd.to_datetime(fy_start), -12, month_begin = True).year
            )
        schedule_start = (
            pd.to_datetime(f'{cutoff_year}-01-01').strftime('%d %B %Y')
            )
        
        # Get the main table
        table_start_row_idx, table_start_col_idx = \
            self.get_indices(df0, 'S/N')[0] #self.get_indices(df0, 'Lease Start')[0]
        
        try:
            table_end_row_idx, table_end_col_idx = \
                self.get_indices(
                    df0, 
                    '<<<END OF TABLE - ROWS AFTER THIS WILL NOT BE READ OR PROCESSED.>>>'
                    )[0]
        except IndexError: 
            msg = ('Template is missing the cell: '
                   '<<<END OF TABLE - ROWS AFTER THIS WILL NOT BE READ OR PROCESSED.>>>'
                   '')
            raise Exception (msg)
            
        # Save attributes
        self.client = client
        self.table_start_row_idx = table_start_row_idx
        self.table_start_col_idx = table_start_col_idx
        self.table_end_row_idx = table_end_row_idx
        self.table_end_col_idx = table_end_col_idx
        self.pfy_start = pfy_start
        self.pfy_end = pfy_end
        self.fy_start = fy_start
        self.fy_end = fy_end
        self.schedule_start = schedule_start
        
    def get_main_data(self):
        
        # Load attributes
        df0 = self.df0.copy()
        
        # Keep only required rows
        table_row_indices = df0.index[
            (df0.index >= self.table_start_row_idx) & 
            (df0.index < self.table_end_row_idx)
            ]
        df = df0.loc[table_row_indices]
        
        # Set the column name and normalise them
        # also, remove the original row with the columns
        df.columns = df.iloc[0].str.strip()
        df = df[df.columns.dropna()].copy()
        df.columns.name = None
        df = df.iloc[1:]
        
        # Identify columns with all NAs
        empty_cols = df.columns[df.isnull().all()]
        # Get the lease groups that are not empty columns
        lease_groups = [c for c in df.columns if (c.startswith("Lease Group") 
                                                  and (c not in empty_cols))]
        
        # Filter only required columns
        mandatory_columns = LeaseDataReader.EXPECTED_COL
                
        # Only keep relevant data
        if lease_groups:
            required_columns = mandatory_columns + lease_groups
        else:
            required_columns = mandatory_columns
        df = df.filter(items = required_columns)
        
        # Check that data contains the required number of columns
        if df.shape[1] != len(required_columns):
            error = "Data table sliced incorrectly"
            logger.exception(error)
            raise Exception (error)
        
        # only keep those with leases
        has_leases = df["Lease Start (PFY)"].notnull() | df["Lease Start"].notnull()
        df_with_leases = df[has_leases]
                
        # Save attributes
        self.df_raw = df_with_leases.copy()
        
    
    def process_data(self):
        
        df = self.df_raw.copy()
        
        def convert_dtypes(input_df, col_lst, op_col_type):
    
            if (op_col_type not in
                ['category',
                 np.number, np.int, 'number', 'integer', 'num', 'int',
                 'str', 'string', 'object', object, np.datetime64]):
                
                raise Exception('Unexpected output column type')
            
            df = input_df.copy()
            
            df[col_lst] = df[col_lst].astype(op_col_type)
            
            return df

        # convert to correct data type
        df = (
            df
            .pipe(convert_dtypes, 
                  ['Lease Start (PFY)', 'Lease End (PFY)', 'Lease Start', 
                   'Lease End', 'Early Termination? (date)'], np.datetime64)
            )
        
        df = (
            df
            .pipe(convert_dtypes,
                  LeaseDataReader.HEADER_NUM_COL, object))
        
        # fill down contract name
        df['Contract'] = df['Contract'].fillna(method="ffill")
        
        # Normalise the string columns
        str_col_lst = \
            df[LeaseDataReader.HEADER_STR_COL].select_dtypes(object).columns
        df[str_col_lst] = (
            df[str_col_lst].apply(lambda x: x.str.strip(), axis=1)
            )
        
        header_cols = LeaseDataReader.HEADER_STR_COL + LeaseDataReader.HEADER_NUM_COL
        header_cols.remove('Renew Contract?')
        header_cols.remove('New Contract')
        header_cols.remove('Early Termination? (date)')
        df[header_cols] = (
            df.groupby('Contract')[header_cols]
            .transform(lambda x: x.ffill().bfill())
            )
        
        df = (
            df
            .pipe(self.generate_contract_type)
            .pipe(self.generate_validation_error_ind)
            .pipe(self.generate_date_information)
            )
        
        df_gb_contract = (df.groupby('Contract')['Validation Error Indicator']
                     .agg(lambda x: x.any()).reset_index())
        if df_gb_contract['Validation Error Indicator'].any():
            check_lst = df_gb_contract.loc[
                df_gb_contract['Validation Error Indicator'], 'Contract'].tolist()
            msg = (
                'Sample input contains incomplete information, please furnish '
                f'the information for contract: {check_lst}.')
            raise Exception(msg)
        
        self.df = df.copy()

    
    
    def generate_contract_type(self, ip_df):
        
        df = ip_df.copy()
        
        target_col = ['New Contract', 'Renew Contract?', 'Early Termination? (date)']
        
        df_gb = df.groupby('Contract', as_index=False).agg({
            'New Contract': 'first',
            'Early Termination? (date)': 'first',
            'Renew Contract?': 'first',
            })
        
        df_tmp = df[set(df.columns) - set (target_col)]
        
        df_tmp = pd.merge(df_tmp, df_gb, on ='Contract')
        
        cond = [
            df_tmp['New Contract'].isna() & df_tmp['Renew Contract?'].isna() & \
                df_tmp['Early Termination? (date)'].isna(),
            df_tmp['New Contract'].notna() & df_tmp['Renew Contract?'].isna() & \
                df_tmp['Early Termination? (date)'].isna(),
            df_tmp['New Contract'].isna() & df_tmp['Renew Contract?'].notna() & \
                df_tmp['Early Termination? (date)'].isna(),
            df_tmp['New Contract'].isna() & df_tmp['Renew Contract?'].isna() & \
                df_tmp['Early Termination? (date)'].notna(),
            df_tmp['New Contract'].isna() & df_tmp['Renew Contract?'].notna() & \
                df_tmp['Early Termination? (date)'].notna(),
            ]
        
        res = [
            'No Change',
            'Addition',
            'Remeasurement',
            'Disposal',
            'Remeasurement + Disposal'
            ]
        
        df['Type'] = np.select(cond, res, 'Unknown')
        
        # Check for unknown contract type
        cond = df['Type'].str.contains('Unknown')
        if cond.any():
            check_lst = list(df.loc[cond, 'Contract'].unique())
            msg = (
                'Input template contains unexpected type of contract. Please '
                f'check {check_lst}.'
                )
            raise Exception(msg)
        
        return df
    
    
    def generate_validation_error_ind(self, ip_df):
        
        df = ip_df.copy()
                
        old_contract_term_col_lst = (
            ['Borrowing Rate (PFY)', 'Rental/mth (PFY)', 
             'Lease Start (PFY)', 'Lease End (PFY)']
            )
        
        new_contract_term_col_lst = (
            ['Borrowing Rate', 'Rental/mth', 'Lease Start', 'Lease End']
            )
                    
        if not (set(df['Type'].unique()) <= 
                set(['No Change', 'Addition', 'Remeasurement', 'Disposal', 
                     'Remeasurement + Disposal'])):
            msg = 'Unexpected Contract Type. Please top up the condition.'
            raise NotImplementedError(msg)
        
        cond = [
            df['Type'].isin(['No Change', 'Disposal']),
            df['Type']=='Addition',
            df['Type'].isin(['Remeasurement', 'Remeasurement + Disposal']),
            ]
        
        res =[
            ~(df[old_contract_term_col_lst].notna().all(axis=1) &
             df[new_contract_term_col_lst].isna().all(axis=1)),
        
            ~(df[old_contract_term_col_lst].isna().all(axis=1) &
             df[new_contract_term_col_lst].notna().all(axis=1)),
            
            ~(df[new_contract_term_col_lst].notna().all(axis=1)),
            
            ]
        
        df['Validation Error Indicator'] = np.select(cond, res, True)
                
        return df
        
    
    def generate_date_information(self, ip_df):
        '''
        Get month column for dataframe
        '''        
        
        df = ip_df.copy()
        
        df['contract_start_date_pfy'] = df.groupby('Contract')['Lease Start (PFY)'].transform('first')
        df['contract_end_date_pfy'] = df.groupby('Contract')['Lease End (PFY)'].transform('last')
        
        df['contract_start_date'] = df.groupby('Contract')['Lease Start'].transform('first')
        df['contract_end_date'] = df.groupby('Contract')['Lease End'].transform('last')
        
        # Update cfy contract date for no change and disposal case
        cond = df['Type'].isin(['No Change', 'Disposal'])
        df.loc[cond, 'contract_start_date'] = df.loc[cond, 'contract_start_date_pfy']
        df.loc[cond, 'contract_end_date'] = df.loc[cond, 'contract_end_date_pfy']
        
        cond = df['Renew Contract?'].isin(['Y'])
        df.loc[cond,'remeasurement_date'] = df.loc[cond,'Lease Start']
        # df['remeasurement_date'] = (
        #     df.groupby('Contract')['remeasurement_date'].transform('first')
        #     )
        df['remeasurement_date'] = (
            df.groupby('Contract')['remeasurement_date']
            .transform(lambda x: x.ffill().bfill()))
        
        df['disposal_date'] = (
            df.groupby('Contract')
            ['Early Termination? (date)'].transform('first'))
        
        df['contract_fy_start'] = (
            [pd.to_datetime(
                datetime.date(date.year,self.fy_start.month, self.fy_start.day)
                )
             for date in 
             df['contract_start_date_pfy'].fillna(df['contract_start_date'])
             ]
            )
        
        df['contract_fy_end'] = (
            [pd.to_datetime(
                datetime.date(date.year, self.fy_end.month, self.fy_end.day)
                )
             for date in 
             df['contract_end_date'].fillna(df['contract_end_date_pfy'])
             ]
            )
        

        return df


#%% Tester
# if __name__ == "__main__":
    
    
#     #%%%% 
#     if 1:
#         # mocked cases
#         input_fp = r"D:\Ben\_Ref\Audit DA Curriculum\Module\frs116_automation\autolease_hm_cw\INPUT TEMPLATE 2021\INPUT TEMPLATE.xlsx"
#         output_fp = r"D:\Ben\_Ref\Audit DA Curriculum\Module\frs116_automation\autolease_hm_cw\INPUT TEMPLATE 2021\OUTPUT.xlsx"
#     if 0:
#         input_fp = r"D:\Ben\_Ref\Audit DA Curriculum\Module\frs116_automation\autolease_hm_cw\INPUT Q&M 2020\INPUT TEMPLATE.xlsx"
#         output_fp  = r"D:\Ben\_Ref\Audit DA Curriculum\Module\frs116_automation\autolease_hm_cw\INPUT Q&M 2020\OUTPUT.xlsx"
    
#     lease_liability_writer = LeaseLiabilityWriter(input_fp, output_fp)
    
#%% Tester
if __name__ == "__main__":
    
    input_fp = r"D:\Ben\_Ref\Audit DA Curriculum\Module\frs116_automation\autolease_hm_cw\INPUT Q&M 2021\INPUT TEMPLATE - FINAL.xlsx"
    input_fp = r"D:\Ben\_Ref\Audit DA Curriculum\Module\frs116_automation\autolease_hm_cw\INPUT Q&M 2021\INPUT TEMPLATE - FINAL edited.xlsx"
    self = LeaseDataReader(input_fp, sheet_name = 'Lease Data',)
    self.__main__()
