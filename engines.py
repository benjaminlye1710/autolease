# -*- coding: utf-8 -*-
"""
Created on Wed Dec 29 18:33:56 2021

@author: AnikaLeeTH
"""
from dependencies import *
from dependencies import ExcelHelper
from pandas import PeriodIndex

# Initialise the excel misc functions to use
Excel_Misc_Fns = ExcelHelper()
XL_Fns = Excel_Misc_Fns

# =============================================================================
#### Engine class
# =============================================================================
class OneContract:
    
    def __init__(self, df, contract, fy_start, fy_end, pfy_start):
        
        self.df_raw = df[df['Contract'] == contract].copy()
        self.contract = contract
        self.type = assert_one_and_get(self.df_raw['Type'].unique())
        self.country = assert_one_and_get(self.df_raw['Country'].unique())
        self.fy_start = fy_start
        self.fy_end = fy_end
        self.pfy_start = pfy_start
        

        # run methods        
        self.get_schedule_start()
        self.get_schedule_end()
        self.prepare_branch_df_month()
        self.prepare_old_branch_df()
        self.prepare_new_branch_df()
            
    def get_schedule_start(self):
        
        """
        Gets starting date of schedule df
        """
        
        df = self.df_raw.copy()
        
        if self.type == 'Remeasurement':
            schedule_start = self.pfy_start
            
        else: # all other cases start from the start of current FY
            dates = []
            frs_start = datetime.datetime(2019,1,1).date()
            dates.append(frs_start)
            contract_fy_start = pd.to_datetime(df['contract_fy_start'].unique()[0])
            
            # adjust month by -12 months -> 1 year ago
            contract_pfy_start = (shift_month(contract_fy_start,
                                              -12, month_begin = True))
            dates.append(contract_pfy_start)
            schedule_start = max(dates) # max finds the more recent date
            
        self.schedule_start = schedule_start
        
        
    def get_schedule_end(self):
        """
        Get ending date of schedule df 
        """
        
        df = self.df_raw.copy()
        
        [schedule_end] = df['contract_fy_end'].dt.date.unique()
        
        self.schedule_end = schedule_end
        
    
    def first_day_month(self, date):
        '''
        to compare if the given date is the first date of the month
        '''
        # replace the day to get the first day of the specified month 
        first_day = date.replace(day=1)
        
        return first_day
    
    
    def last_day_month(self, date):
        '''
        to compare if the given date is the first date of the month
        '''
        
        # use calendar package to get the max number of days of the specified month and year
        date_num = calendar.monthrange(date.year, date.month)[1]
        
        # replace the day to get the last day of the specified month and year
        last_day = date.replace(day=date_num)
        
        return last_day
        
        
    def prepare_branch_df_month(self):
        
        df = self.df_raw.copy()
        cond = df['Type'].str.contains('Remeasurement')
        cond_add = df['Type'].str.contains('Addition')
        
        branch_df = pd.DataFrame()
        
        # fill in branch_df with "Month" column
        if self.type=='Remeasurement':
            
            if self.schedule_end < self.fy_start:
                branch_df['Month'] = (
                    pd.period_range(self.schedule_start, self.fy_end, freq = 'M')
                    )
            else:
                # Month column in excel, returns a fixed period index
                branch_df['Month'] = (
                    pd.period_range(self.schedule_start, self.schedule_end, freq = 'M')
                    )

        else:
            agg_dict = {
                'contract_fy_start': 'first',
                'contract_fy_end': 'first',}

            branch_df = df.groupby('Contract').agg(agg_dict)
            
            full_date_range_col_lst = ['contract_fy_start',
                                       'contract_fy_end',]
            
            if self.schedule_end < self.fy_start:
                branch_df['Month'] = (
                    branch_df[full_date_range_col_lst]
                    .apply(lambda x: pd.period_range(x[0], self.fy_end, freq = 'M'), 
                           axis = 1)
                    )
            else:
                branch_df['Month'] = (
                    branch_df[full_date_range_col_lst]
                    .apply(lambda x: pd.period_range(x[0], x[1], freq = 'M'), 
                           axis = 1)
                    )
                
            # explode() function is used to transform each element of a list-like to a row
            # replicating the index values
            branch_df = branch_df.explode('Month')
        
        df_unpivot = df.copy()
        
        # Calculate Fixed Lease Payment without pro-rating yet
        df_unpivot['Fixed lease payment'] = df_unpivot['Rental/mth'].copy()
        df_unpivot['Fixed lease payment (PFY)'] = df_unpivot['Rental/mth (PFY)'].copy()

        # remeasurement/addition case
        if self.type in ['Remeasurement', 'Addition']:
            
            start = 'Lease Start'
            end = 'Lease End'
            start_mth = 'Lease Start Month'
            end_mth = 'Lease End Month'
            start_day = 'Start Day'
            end_day = 'End Day'
            rental = 'Fixed lease payment'
            rental_l1 = 'Fixed lease payment l1'
            rental_2 = 'Fixed lease payment (PFY)'
            rental_2_l1 = 'Fixed lease payment (PFY) l1'
        
        else:
            
            start = 'Lease Start (PFY)'
            end = 'Lease End (PFY)'
            start_mth = 'Lease Start Month (PFY)'
            end_mth = 'Lease End Month (PFY)'
            start_day = 'Start Day (PFY)'
            end_day = 'End Day (PFY)'
            rental = 'Fixed lease payment (PFY)'
            rental_l1 = 'Fixed lease payment (PFY) l1'
            rental_2 = 'Fixed lease payment'
            rental_2_l1 = 'Fixed lease payment l1'

            
        ## Copy Lease Start and Lease End Months for each contract before exploding all the months
        df_unpivot = df_unpivot.reset_index(drop=True)
        df_unpivot[start_mth] = df_unpivot[start].dt.to_period(freq = ' M') ##
        df_unpivot[end_mth] = df_unpivot[end].dt.to_period(freq = ' M') ##
        
        ## Add Month column with full range of months based on Lease Start and Lease End columns
        # df_unpivot['Month'] = (
        #     df_unpivot[[start, end]]
        #     .apply(lambda x: pd.date_range(x[0], x[1], freq = 'M'), 
        #             axis = 1)
        #     )
        
        df_unpivot['Month'] = (
            df_unpivot[[start_mth, end_mth]]
            .apply(lambda x: pd.period_range(x[0], x[1], freq = 'M'), 
                    axis = 1)
            )
        
        df_unpivot = df_unpivot.explode('Month') # Populating all the months into individual rows
        
        ## remove duplicated rows"
        df_unpivot['lead'] = df_unpivot['Month'].shift(-1)

        duplicated_rows = df_unpivot['Month'] == df_unpivot['lead']
        df_unpivot = df_unpivot.loc[~duplicated_rows, :].copy()
        
        ## adjust prorate days for factor calculation
        df_unpivot[start_day] = df_unpivot[start].dt.day
        df_unpivot[end_day] = df_unpivot[end].dt.day

        df_unpivot['Total Days in Month'] = df_unpivot['Month'].apply(
            lambda x: calendar.monthrange(x.year, x.month)[1]
            )
        df_unpivot['Ending Day in Month'] = df_unpivot['Total Days in Month']
        
        df_unpivot['lag'] = df_unpivot[end_mth].shift(1)
        df_unpivot[rental_l1] = df_unpivot[rental].shift()
        df_unpivot[rental_2_l1] = df_unpivot[rental_2].shift()

        if self.type in ['No Change', 'Disposal']:
            df_unpivot[rental_2] = df_unpivot[rental] 
            df_unpivot[rental_2_l1] = df_unpivot[rental_l1]
        
        ## if contract does not start on the first date
        not_first_day = (df_unpivot[start_day] != 1).any()
        if not_first_day:
            df_unpivot[rental_l1].iat[0] = 0
            df_unpivot[rental_2_l1].iat[0] = 0
        
        ## if contract end date is not 1 day before start date
        lease_mth_end_day = (
            df_unpivot[end].apply(lambda x: calendar.monthrange(x.year, x.month)[1])
            )
        not_last_day = (df_unpivot[end_day] != lease_mth_end_day).any()
        consecutive_start_and_end = ((df_unpivot[start_day]-1) == df_unpivot[end_day]).all()
        
        if not_last_day & consecutive_start_and_end:
            df_unpivot[rental].iat[-1] = 0
            df_unpivot[rental_2].iat[-1] = 0
            
        elif (not_last_day & (df_unpivot[start_day] < df_unpivot[end_day]).all()):
            df_unpivot[end_day] = df_unpivot[start_day]-1
            df_unpivot['Ending Day in Month'].iat[-1] = df_unpivot[end].dt.day.iat[-1]
        
        elif ~not_last_day:
            df_unpivot[end_day] = 0
            
        else:
            raise Exception
        
        df_unpivot = df_unpivot.reset_index(drop=True) # reset index
        
        # Assuming rental is calculated from 16th this month till 15th next month
        # Assuming rental per month increase from 
        # $1000 (16 Jan - 15 Feb) to $1200 (16 Feb - 15 Mar)
        # this formula will calculate the rental in Feb
        # 1st part of rental 1 Feb 2021 - 15 Feb 2021 using previous month's rental $1000
        df_unpivot['factor_l1'] = df_unpivot[end_day] / df_unpivot['Total Days in Month']
        # 2nd part of rental 16 Feb 2021 - 28 Feb 2021 using current month's rental $1200
        df_unpivot['factor'] = (
            (df_unpivot['Ending Day in Month'] - df_unpivot[start_day] + 1) /
            df_unpivot['Total Days in Month']
            )
        
        df_unpivot[rental] = (
            df_unpivot[rental_l1].fillna(0) * df_unpivot['factor_l1'] +
            df_unpivot[rental] * df_unpivot['factor']
            )
                
        df_unpivot[rental_2] = (
            df_unpivot[rental_2_l1].fillna(0) * df_unpivot['factor_l1'] +
            df_unpivot[rental_2] * df_unpivot['factor']
            )
            
        # test = df_unpivot[
        #     ['Month', start, start_mth, 
        #      # 'not_start', 
        #       end, end_mth, 'lag', 
        #       # 'not_end',
        #       start_day, end_day, 'Ending Day in Month', 'Total Days in Month',
        #       'factor', 'factor_l1', 
        #       rental, rental_l1, rental_2, rental_2_l1,
        #       ]
        #     ].copy()
                
        self.df = df.copy()
        self.branch_df = branch_df.copy()
        self.df_unpivot = df_unpivot.copy()
        
    def prepare_old_branch_df(self):

        df = self.df.copy()
        branch_df = self.branch_df.copy()
        df_unpivot = self.df_unpivot.copy()
        
        key_id = ['Contract', 'Month','Fixed lease payment (PFY)']
        #agg_dict = {'Fixed lease payment': 'min'}
        df_unpivot_2 = df_unpivot.loc[:, key_id].copy()
        
        #df_unpivot['RowID'] = range(len(df_unpivot))
        
        branch_df['Month'] = branch_df['Month'].dt.strftime('%b-%y')
        df_unpivot_2['Month'] =  (
            pd.to_datetime(df_unpivot_2['Month'].astype(str))
            .dt.strftime('%b-%y').copy())

        if self.type=='Remeasurement':
            try:
                branch_df = (branch_df.merge(df_unpivot_2, on = ['Month'],
                             how = 'left', validate = '1:1'))
                branch_df['Contract'] = branch_df['Contract'].fillna(method='ffill')
            except Exception:
                msg = 'ERROR'
                raise Exception(msg)
        else:
            try: 
                branch_df = (branch_df.merge(df_unpivot_2, on = ['Contract','Month'],
                                             how = 'left', validate = '1:1'))
            except Exception:
                msg = 'ERROR'
                raise Exception(msg)
        
        branch_df[''] = (branch_df.groupby('Contract')['Contract']
                         .cumcount() + 1)
        
        col_filter_lst = ['Contract','Month','','Fixed lease payment (PFY)']
        branch_df = branch_df[col_filter_lst].copy()
        
        num_months_df = (branch_df.groupby('Contract').agg(
            **{'months_from_fystart':('Month','count')}).rename(
                columns = {'Contract':'branch'}))
        
        self.old_branch_df = branch_df.copy()
        self.old_df_unpivot_2 = df_unpivot_2.copy()
        self.old_num_months_df = num_months_df.copy()
        
    def prepare_new_branch_df(self):
        """
        Prepare new branch sheet 
        """
        
        df = self.df.copy()
        branch_df = self.branch_df.copy()
        df_unpivot = self.df_unpivot.copy()
        
        key_id = ['Contract', 'Month','Fixed lease payment']
        df_unpivot_2 = df_unpivot.loc[:, key_id].copy()
        
        branch_df['Month'] = branch_df['Month'].dt.strftime('%b-%y')
        df_unpivot_2['Month'] = (
            pd.to_datetime(df_unpivot_2['Month'].astype(str)).dt.strftime('%b-%y')
            )
        
        if self.type=='Remeasurement':
            try:
                branch_df = (branch_df.merge(df_unpivot_2, on = ['Month'],
                             how = 'left', validate = '1:1'))
                branch_df['Contract'] = branch_df['Contract'].fillna(method='ffill')
            except Exception:
                msg = 'ERROR'
                raise Exception(msg)
        else:
            try: 
                branch_df = (branch_df.merge(df_unpivot_2, on = ['Contract','Month'],
                                             how = 'left', validate = '1:1'))
            except Exception:
                msg = 'ERROR'
                raise Exception(msg)
            
        branch_df[''] = (branch_df.groupby('Contract')['Contract']
                         .cumcount() + 1)
        
        col_filter_lst = ['Contract','Month','','Fixed lease payment']
        branch_df = branch_df[col_filter_lst].copy()
                    
        num_months_df = (branch_df.groupby('Contract').agg(
            **{'months_from_fystart':('Month','count')}).rename(
                columns = {'Contract':'branch'}))
        
        self.new_branch_df = branch_df.copy()
        self.new_df_unpivot_2 = df_unpivot_2.copy()
        self.new_num_months_df = num_months_df.copy()
        

        
        
class OneContractSchedule(OneContract):
    
    def __init__(self, df, contract, fy_start, fy_end, pfy_start, output_fp,):
        
        OneContract.__init__(self, df, contract, fy_start, fy_end, pfy_start,)
        
        self.fp = output_fp
        wb = openpyxl.load_workbook(output_fp)
        ws = wb.copy_worksheet(wb['branch_template'])
        
        old_branch_df = self.old_branch_df.copy()
        new_branch_df = self.new_branch_df.copy()
        
        non_empty_row = old_branch_df['Fixed lease payment (PFY)'].notnull()
        old_branch_df_not_null = old_branch_df.loc[non_empty_row, :].copy()
        
        non_empty_row = new_branch_df['Fixed lease payment'].notnull()
        new_branch_df_not_null = new_branch_df.loc[non_empty_row, :].copy()
        
        # get first mth & year thats not NA
        try:
            old_first_m_y = old_branch_df_not_null.iat[0, 1]
        except Exception:
            old_first_m_y = old_branch_df.iat[0, 1]
            
        new_first_m_y = new_branch_df_not_null.iat[0, 1]

        # get num of months & years
        try:
            old_num_months = int(old_branch_df_not_null[""].max())
            old_num_years = math.ceil(old_num_months/12)
        except Exception:
            old_num_months = 0
            old_num_years = 0

        new_num_months = new_branch_df_not_null[""].max()
        # new_num_months = new_branch_df[""].max()
        new_num_years = math.ceil(new_num_months/12)
        # new_num_years = math.ceil(new_branch_df[""].max()/12)
        
        # get first row start
        first_row = 8

        first_month = old_branch_df['Month']==old_first_m_y
        old_lease_start_row = old_branch_df.loc[first_month].index[0]
        old_lease_start_row_idx = old_lease_start_row + first_row
                
        first_month = new_branch_df['Month']==new_first_m_y
        new_lease_start_row = new_branch_df.loc[first_month].index[0]
        new_lease_start_row_idx = new_lease_start_row + first_row
        
        # save
        self.wb = wb
        self.ws = ws
        self.first_row = first_row
        
        self.old_branch_df_not_null = old_branch_df_not_null
        self.old_first_m_y = old_first_m_y
        self.old_num_months = old_num_months
        self.old_num_years = old_num_years
        self.old_lease_start_row = old_lease_start_row
        self.old_lease_start_row_idx = old_lease_start_row_idx
        self.old_final_row_idx = str(self.first_row + old_num_months - 1)
        self.old_final_row_idx_full = (
            str(self.first_row + (old_num_years*12)-1)
            )
            # str(len(self.old_branch_df))
            
        self.new_branch_df_not_null = new_branch_df_not_null
        self.new_first_m_y = new_first_m_y
        self.new_num_months = new_num_months
        self.new_num_years = new_num_years
        self.new_lease_start_row = new_lease_start_row
        self.new_lease_start_row_idx = new_lease_start_row_idx
        self.new_final_row_idx = str(self.first_row + new_num_months - 1)
        self.new_final_row_idx_full = (
            str(self.first_row + (new_num_years*12)-1)
            )
            # str(len(self.new_branch_df))
            
        
    def __main__(self):
        
        '''
        To call all the other functions 
        '''

        old_schedule_df = self.get_old_schedule_df().copy()
        new_schedule_df = self.get_new_schedule_df().copy()
        
        self.schedule_df = pd.concat([old_schedule_df,new_schedule_df],
                                axis=1)
                
        # Converting column names to match excel's columns        
        cols = list('ABCDEFGHIJKLMNOPQRSTUVWXYZ')
        self.schedule_df.columns = cols[:len(self.schedule_df.columns)]

        self.write_meta_data()
        self.__insert_rows__()
        self.write_old_interest_rate()
        self.write_new_interest_rate()
        self.write_total_lease_period()
        self.write_schedule_df()
        self.write_old_period_formula()
        self.write_new_period_formula()
        
        if self.type in ['Remeasurement','Disposal']:
            try:
                self.get_disposal_row_idx(self.old_lease_payment_df)
            except Exception:
                pass
            self.conditional_format()
            
        
    def __calculate_final_row_indices(self,num_months,num_years):
        
        # get final row indices
        final_row_idx = str(self.first_row + num_months - 1)
        final_row_idx_full = str(self.first_row + (num_years*12)-1)
        
        return final_row_idx, final_row_idx_full
        
    ## to fill the dates column
    def fill_full_year(self, branch_df):
        
        branch_df = branch_df.copy()
        
        if branch_df.shape[0] % 12 != 0:
            for i in np.arange(12-(len(branch_df)%12)):
                branch_df.loc[branch_df.iloc[-1].name + i + 1, :] = np.nan
        
        return branch_df
    
    def get_old_df_summary(self):
        
        '''
        Forms the old column summary (last line)
        '''
        
        branch_df = self.old_branch_df.copy()
        first_row = self.first_row
        final_row_idx_full = self.old_final_row_idx_full

        # expand & fill full year
        full_year_df = self.fill_full_year(branch_df)
        
        # get start & end cell
        start_cell = Excel_Misc_Fns.xlref(first_row,3)
        end_cell = Excel_Misc_Fns.xlref(final_row_idx_full,3)
        
        # sum for fixed lease payment
        flp_total = Excel_Misc_Fns.sum_formula(start_cell,end_cell)
        
        # append 'Total' as new row
        row = pd.Series({'Month'                :'Total',
                         ''                     :'',
                         'Fixed lease payment (PFY)'  :flp_total})

        cond = self.df['Type'].str.contains('Addition')
        if cond.all(): # does not append total row if addition case (empty df)
            df_summed = full_year_df
        else:
            df_summed = full_year_df.append([row], ignore_index = True)
        
        add_cols = ['Lease Incentives (PFY)', 'Additional Payments (PFY)',
                    'Variable Payments (PFY)']
        
        for add_col in add_cols:
            df_summed[add_col] = ''
        
        self.old_lease_payment_df = df_summed.copy()
        
        return df_summed
    
    
    def get_new_df_summary(self):
        
        '''
        Forms the new column summary (last line)
        '''
        
        branch_df = self.new_branch_df.copy()
        first_row = self.first_row
        final_row_idx_full = self.new_final_row_idx_full

        # expand & fill full year
        full_year_df = self.fill_full_year(branch_df)
        
        # get start & end cell
        start_cell = Excel_Misc_Fns.xlref(first_row,12)
        end_cell = Excel_Misc_Fns.xlref(final_row_idx_full,12)
        
        # sum for fixed lease payment
        flp_total = Excel_Misc_Fns.sum_formula(start_cell,end_cell)
        
        # append 'Total' as new row
        row = pd.Series({'Month'                :'Total',
                         ''                     :'',
                         'Fixed lease payment'  :flp_total})
        # row.name = int(self.new_final_row_idx)-1
        row.name = full_year_df.index.max() + 1
        df_summed = full_year_df.append([row])
        
        cond = self.df['Type'].str.contains('Remeasurement')
        if cond.all():  
            self.remeasurement_update_df(df_summed)
            
        add_cols = ['Lease Incentives', 'Additional Payments',
                    'Variable Payments']
        
        for add_col in add_cols:
            df_summed[add_col] = ''
        
        self.new_lease_payment_df = df_summed.copy()
        
        return df_summed
    
    def remeasurement_update_df(self,df_summed):
        
        """
        Updates flp for the month of remeasurement -> finds the row idx of the remeasurement
        """
        
        df = self.df.copy()
        first_row = self.first_row
        
        remeasurement_dates = \
            df['remeasurement_date'].dt.strftime('%b-%y').unique()
            
        remeasurement_row = \
            df_summed[df_summed['Month'] == remeasurement_dates[-1]].index[0] 
        remeasurement_row_idx = remeasurement_row + first_row
        
        remeasurement_days = pd.to_datetime(
            df['remeasurement_date'].unique()[-1]).day
        
        prev_flp = df_summed.loc[remeasurement_row_idx-first_row-1,'Fixed lease payment']
        new_flp = df_summed.loc[remeasurement_row_idx-first_row,'Fixed lease payment']
        
        df_summed.loc[remeasurement_row_idx-first_row,'Fixed lease payment'] = \
            (prev_flp/31)*(remeasurement_days-1) + (new_flp/31)*(31-remeasurement_days+1)
        
        self.remeasurement_row_idx = remeasurement_row_idx
        
    def get_disposal_row_idx(self,df_summed):
        
        '''
        Gets the row idx of disposal date
        '''
        
        df = self.df.copy()
        first_row = self.first_row
        
        disposal_date = df['disposal_date'].dt.strftime('%b-%y').unique()
        disposal_row = \
            df_summed[df_summed['Month'] == disposal_date[0]].index[0]
        disposal_row_idx = disposal_row + first_row
        
        self.disposal_row_idx = disposal_row_idx
        
    
    def get_schedule_period_start(self):
        """
        Gets the starting number for period column
        """
        df = self.df.copy()
        schedule_start = pd.to_datetime(self.schedule_start)
        
        # adjust month by -12 months -> 1 year ago
        ppfy_start = (shift_month(pd.to_datetime(self.schedule_start),
                                  -12, month_begin = True))
        
        remeasurement_dates = df['remeasurement_date'].unique()
        
        dates_before_pfy_start = []
        for x in remeasurement_dates:
            if pd.to_datetime(x) < schedule_start: # if remeasurement date is earlier than sched start
                # add remeasurement date to "dates_before_pfy_start" list
                dates_before_pfy_start.append(pd.to_datetime(x))
                
        if not dates_before_pfy_start: # if list is empty which means all remeasurement dates are later than pfy start
            # get contract start date
            contract_start = pd.to_datetime(
                df['contract_start_date_pfy'].unique()[0])
            # number of months from contract start date
            schedule_period_start = (schedule_start.year - contract_start.year)*12 \
                + (schedule_start.month - contract_start.month) + 1
                
            # if schedule start is earlier than contract start, return 0
            schedule_period_start = max(0,schedule_period_start) ###edited from 0 to 1
        else:
            last_remeasurement = dates_before_pfy_start[-1]
            schedule_period_start = (schedule_start.year - last_remeasurement.year)*12 \
                + (schedule_start.month - last_remeasurement.month) + 1
                
        return schedule_period_start 
    
    def old_cell_ref(self, row_idx):
        """
        Generates a dataframe with cells to be filled with formulae 
        and their cell reference through generating a dictionary 
        with cell references required.
        """
        
        row_idx = int(row_idx)
        cells = {
            'pfy_flp'           : Excel_Misc_Fns.xlref(row_idx, 3),
            'pfy_vlp'           : Excel_Misc_Fns.xlref(row_idx, 6),
                
            # columns to fill start (idx = 2)
            'pfy_period'        : Excel_Misc_Fns.xlref(row_idx, 2),
            'pfy_tfp'           : Excel_Misc_Fns.xlref(row_idx, 7),
            'pfy_int_rate'      : Excel_Misc_Fns.xlref(row_idx, 8),
            'pfy_NPV'           : Excel_Misc_Fns.xlref(row_idx, 9),
            
            # columns to fill end (idx = 4)
            
            'pfy_period_lag'    : Excel_Misc_Fns.xlref(row_idx-1, 2)

            }
        
        self.old_cells = cells
        
        old_cells_df = pd.Series(cells).to_frame(name="cell_ref").iloc[2:6]
        
        return old_cells_df
        
    def new_cell_ref(self,row_idx):
        """
        Generates a dataframe with cells to be filled with formulae 
        and their cell reference through generating a dictionary 
        with cell references required.
        """
        
        row_idx = int(row_idx)
        cells = {
            'flp'           : Excel_Misc_Fns.xlref(row_idx, 12),
            'vlp'           : Excel_Misc_Fns.xlref(row_idx, 15),
                
            # columns to fill start (idx = 2)
            'period'        : Excel_Misc_Fns.xlref(row_idx, 11),
            'tfp'           : Excel_Misc_Fns.xlref(row_idx, 16),
            'int_rate'      : Excel_Misc_Fns.xlref(row_idx, 17),
            'NPV'           : Excel_Misc_Fns.xlref(row_idx, 18),
            
            'll'            : Excel_Misc_Fns.xlref(row_idx, 19),
            'll_payment'    : Excel_Misc_Fns.xlref(row_idx, 20),
            'll_int'        : Excel_Misc_Fns.xlref(row_idx, 21),
            'CF_ll'         : Excel_Misc_Fns.xlref(row_idx, 22),
                
            'la'            : Excel_Misc_Fns.xlref(row_idx, 23),
            'dep'           : Excel_Misc_Fns.xlref(row_idx, 24),
            'CF_la'         : Excel_Misc_Fns.xlref(row_idx, 25),
            # columns to fill end (idx = 12)
                
            # to calculate columns
            'period_lag'    : Excel_Misc_Fns.xlref(row_idx-1, 11),
            'll_lag'        : Excel_Misc_Fns.xlref(row_idx-1, 19),
            'CF_ll_lag'     : Excel_Misc_Fns.xlref(row_idx-1, 22),
            'la_lag'        : Excel_Misc_Fns.xlref(row_idx-1, 23),
            'dep_lag'       : Excel_Misc_Fns.xlref(row_idx-1, 24),
            'CF_la_lag'     : Excel_Misc_Fns.xlref(row_idx-1, 25)
            }
        
        self.new_cells = cells
        
        new_cells_df = pd.Series(cells).to_frame(name="cell_ref").iloc[2:13]
        
        return new_cells_df
        
        
    def get_old_formulae(self,cells_df,row_idx):
        
        '''
        Stating the cell ref for EACH row idx
        '''
        
        cells = self.old_cells
        first_row = self.first_row
        final_row_idx_full = self.old_final_row_idx_full
        lease_start_row_idx = self.old_lease_start_row_idx
        
        # get excel fn
        XL_Fns = Excel_Misc_Fns
        sum_formula = XL_Fns.sum_formula
        addition_formula = XL_Fns.addition_formula
        subtraction_formula = XL_Fns.subtraction_formula
        divide_formula = XL_Fns.divide_formula
        if_formula = XL_Fns.if_formula
        isblank_formula = XL_Fns.isblank_formula
        xlref = XL_Fns.xlref
        
        one_mth_int = '($S$1/12)'
        total_period = 'M4'
        row_no_from_lease_start = row_idx - lease_start_row_idx
        int_period_ref = row_no_from_lease_start if row_no_from_lease_start > 0 else 0
        
        cells_df.at['pfy_period', 'cell_value'] = if_formula(
            isblank_formula(cells['pfy_flp'], equal_sign=False),
            0,
            addition_formula([cells['pfy_period_lag'],1], equal_sign=False))
        cells_df.at['pfy_tfp', 'cell_value'] = \
            sum_formula(cells['pfy_flp'], cells['pfy_vlp'])
            
        cells_df.at['pfy_int_rate', 'cell_value'] = \
            addition_formula(['1', one_mth_int]) + '^(' + subtraction_formula(
                [cells['pfy_period'], '1'], equal_sign = False) + ')'
            
        cells_df.at['pfy_NPV', 'cell_value'] = \
            divide_formula(cells['pfy_tfp'], cells['pfy_int_rate'])
            
        cells_df_filled = cells_df
        
        return cells_df_filled
    
    
    def get_new_formulae(self,cells_df,row_idx):
        
        '''
        Stating the cell ref for EACH row idx
        '''
        
        cells = self.new_cells
        first_row = self.first_row
        final_row_idx_full = self.new_final_row_idx_full
        lease_start_row_idx = self.new_lease_start_row_idx
        
        df = self.df
        
        # get excel fn
        XL_Fns = Excel_Misc_Fns
        sum_formula = XL_Fns.sum_formula
        addition_formula = XL_Fns.addition_formula
        subtraction_formula = XL_Fns.subtraction_formula
        divide_formula = XL_Fns.divide_formula
        if_formula = XL_Fns.if_formula
        isblank_formula = XL_Fns.isblank_formula
        xlref = XL_Fns.xlref
        
        #old_one_mth_int = '($S$1/12)'
        new_one_mth_int = '($S$2/12)'
        
        #old_total_period = 'M4'
        new_total_period = 'N4'
        
        row_no_from_lease_start = row_idx - lease_start_row_idx
        int_period_ref = row_no_from_lease_start if row_no_from_lease_start > 0 else 0
        
        cells_df.at['period', 'cell_value'] = if_formula(
            isblank_formula(cells['flp'], equal_sign=False),
            0,
            addition_formula([cells['period_lag'],1], equal_sign = False))
        
        cells_df.at['tfp', 'cell_value'] = \
            sum_formula(cells['flp'], cells['vlp'])
            
        cells_df.at['int_rate', 'cell_value'] = \
            addition_formula(['1', new_one_mth_int]) + '^(' + subtraction_formula(
                [cells['period'], '1'], equal_sign = False) + ')'
            
        cells_df.at['NPV', 'cell_value'] = \
            divide_formula(cells['tfp'], cells['int_rate'])
        
        cells_df.at['ll_payment', 'cell_value'] = '=-' + cells['tfp']
        
        cells_df.at['ll_int', 'cell_value'] = \
            sum_formula(cells['ll'], cells['ll_payment']) + '*' + new_one_mth_int
            
        cells_df.at['CF_ll', 'cell_value'] = \
            sum_formula(cells['ll'], cells['ll_int'])
        
        cells_df.at['CF_la', 'cell_value'] = \
            sum_formula(cells['la'], cells['dep'])
        
        if row_idx == first_row:
            
            contract_type = assert_one_and_get(df['Type'].unique())
            if contract_type == 'Addition':
                cells_df.at['ll', 'cell_value'] = \
                    f"=IF(ISBLANK({cells['flp']}),0,SUM(R{first_row}:R{final_row_idx_full}))"
            else:
                cells_df.at['ll', 'cell_value'] = \
                    f"=IF(ISBLANK({cells['flp']}),0,SUM(I{first_row}:I{final_row_idx_full}))"
            
            cells_df.at['la', 'cell_value'] = '=' + cells['ll']
            
            cells_df.at['dep', 'cell_value'] = divide_formula(
                xlref(row_idx, 23, negative=True), 
                new_total_period)
        
        else:
            
            cells_df.at['ll', 'cell_value'] = if_formula(
                'AND(' + cells['ll_lag'] + '=0,' + cells['la_lag'] + '=0)', 
                if_formula(
                    isblank_formula(cells['flp'], equal_sign=False), 
                    0, 
                    sum_formula(xlref(first_row, 9), 
                                xlref(final_row_idx_full, 9), 
                                equal_sign=False), 
                    equal_sign=False), 
                cells['CF_ll_lag']
                )
            
            cells_df.at['la', 'cell_value'] = if_formula(
                cells['la_lag'] + '=0', 
                cells['ll'], 
                cells['CF_la_lag']
                )
            
            cells_df.at['dep', 'cell_value'] = if_formula(
                cells['dep_lag'] + '=0', 
                divide_formula('-' + cells['la'], 
                               new_total_period, 
                               equal_sign=False), 
                'SUMIF($' + cells['CF_la_lag'] + ',">1",' + cells['dep_lag'] + ')'
                )
            
            cond = self.df['Type'].isin(['Addition'])
            if cond.all():
                cells_df.at['ll', 'cell_value'] = if_formula(
                    'AND(' + cells['ll_lag'] + '=0,' + cells['la_lag'] + '=0)', 
                    if_formula(
                        isblank_formula(cells['flp'], equal_sign=False), 
                        0, 
                        sum_formula(xlref(first_row, 18), 
                                    xlref(final_row_idx_full, 18), 
                                    equal_sign=False), 
                        equal_sign=False), 
                    cells['CF_ll_lag']
                    )
        
        cells_df_filled = cells_df
        
        return cells_df_filled
    
    def get_old_formulae_df(self):
        
        '''
        Forms the df by looping every row inside this fn to populate the cell ref/formula
        '''
        
        final_row_idx_full = self.old_final_row_idx_full
        first_filled_row = self.old_lease_start_row_idx
        num_rows = self.old_num_months 
        first_row = self.first_row
        
        
        XL_Fns = Excel_Misc_Fns
        sum_formula = XL_Fns.sum_formula
        xlref = XL_Fns.xlref 
        
        row_formulae = []
        for num in np.arange(num_rows):
            row_idx = first_row + num
            cells_df = self.old_cell_ref(row_idx)
            cells_df_filled = self.get_old_formulae(cells_df, row_idx)
            
            row = cells_df_filled.loc[:, 'cell_value'].tolist()
            row_formulae.append(row)
            
        row_df = pd.DataFrame(row_formulae).reset_index(drop=True)
        row_df = row_df.reindex(self.old_branch_df.index).copy()
        
        full_year_df = self.fill_full_year(row_df)
        
        # Get total row formulae
        cols_alpha = {'pfy_tfp_total': 7, 'pfy_NPV_total': 9}
        total = {}
        for col_name, col_num in cols_alpha.items():
            formulae = sum_formula(xlref(first_row, col_num), 
                                   xlref(final_row_idx_full, col_num))
            total[col_name] = formulae
        
        cond = self.df['Type'].str.contains('Addition')
        if cond.all(): # addition case -> generates empty columns 
            total_row = ['', '', '', '']
        else:
            total_row = ['',total['pfy_tfp_total'], '', total['pfy_NPV_total']]
            
        formulae_df = full_year_df.append([total_row], ignore_index=True)
        
        self.old_main_table_formulae_df = formulae_df.copy()
        
        return formulae_df
    
    def get_new_formulae_df(self):
        
        '''
        Forms the df by looping every row inside this fn to populate the cell ref/formula
        '''
        
        # Get attributes
        final_row_idx_full = self.new_final_row_idx_full # Excel last row 
        first_filled_row = self.new_lease_start_row_idx # Excel first lease row
        num_rows = self.new_num_months # total count from first lease row
        first_row = self.first_row # Excel first fy month row
        
        # Get excel fn
        XL_Fns = Excel_Misc_Fns
        sum_formula = XL_Fns.sum_formula
        xlref = XL_Fns.xlref
        
        row_formulae = []
        for num in np.arange(num_rows):
            row_idx = first_row + num
            cells_df = self.new_cell_ref(row_idx)
            cells_df_filled = self.get_new_formulae(cells_df, row_idx)
            if self.type=='Remeasurement':
                cells_df_filled = self.remeasurement_update_formulae(cells_df_filled,row_idx)
            
            row = cells_df_filled.loc[:, 'cell_value'].tolist()
            row_formulae.append(row)
            
        row_df = pd.DataFrame(row_formulae).reset_index(drop=True)
        row_df = row_df.reindex(self.new_branch_df.index).copy()
        
        full_year_df = self.fill_full_year(row_df).copy()
        
        # Get total row formulae
        cols_alpha = {'tfp_total': 16, 'NPV_total': 18, 'll_payment_total': 20, 
                      'll_int_total': 21, 'dep_charge_total': 24}
        total = {}
        for col_name, col_num in cols_alpha.items():
            formulae = sum_formula(xlref(first_row, col_num), 
                                   xlref(final_row_idx_full, col_num))
            total[col_name] = formulae
        
        total_row = ['',total['tfp_total'], '', total['NPV_total'], 
                     '', total['ll_payment_total'], 
                     total['ll_int_total'], '', '', total['dep_charge_total'],'']
        formulae_df = full_year_df.append([total_row], ignore_index=True)
        
        self.new_main_table_formulae_df = formulae_df.copy()
        
        return formulae_df
    
    def remeasurement_update_formulae(self, cells_df_filled, row_idx):
        """
        Updates the row of formulas which remeasurement occurs 
        """
        
        XL_Fns = Excel_Misc_Fns
        sum_formula = XL_Fns.sum_formula
        addition_formula = XL_Fns.addition_formula
        subtraction_formula = XL_Fns.subtraction_formula
        divide_formula = XL_Fns.divide_formula
        if_formula = XL_Fns.if_formula
        isblank_formula = XL_Fns.isblank_formula
        xlref = XL_Fns.xlref
        
        cells = self.new_cells
        df = self.df.copy()
        new_branch_df = self.get_new_df_summary()
        first_row = self.first_row
        
        new_one_mth_int = '($S$2/12)'
        old_one_mth_int = '($S$1/12)'
        
        old_total_period = 'M4'
        new_total_period = 'N4'
        ll_remeasurement = 'AD40'
        
        remeasurement_row_idx = self.remeasurement_row_idx
        
        if row_idx == remeasurement_row_idx:
            cells_df_filled.at['period', 'cell_value'] = 1
            cells_df_filled.at['int_rate', 'cell_value'] = \
                addition_formula(['1', new_one_mth_int]) + '^(' + subtraction_formula(
                [cells['period'], '1'], equal_sign = False) + ')'
            
            cells_df_filled.at['ll_int', 'cell_value'] = \
                addition_formula([cells['ll'], cells['ll_payment'], \
                    ll_remeasurement]) \
                        + '*' + new_one_mth_int
            cells_df_filled.at['CF_ll', 'cell_value'] = \
                sum_formula(cells['ll'], cells['ll_int']) + '+' + ll_remeasurement
            cells_df_filled.at['dep','cell_value'] = \
                divide_formula('(' + '-' + cells['la'] + '-' + ll_remeasurement + ')',
                                new_total_period)    
            cells_df_filled.at['CF_la', 'cell_value'] = \
                sum_formula(cells['la'], cells['dep']) + "+" + ll_remeasurement
        
        if (row_idx < remeasurement_row_idx and row_idx != first_row):
            cells_df_filled.at['int_rate', 'cell_value'] = \
                addition_formula(['1', old_one_mth_int]) + '^(' + subtraction_formula(
                [cells['period'],'1'], equal_sign = False) + ')'
            cells_df_filled.at['ll_int', 'cell_value'] = \
                sum_formula(cells['ll'], cells['ll_payment']) + '*' + old_one_mth_int
            cells_df_filled.at['dep', 'cell_value'] = if_formula(
                cells['dep_lag'] + '=0', 
                divide_formula('-' + cells['la'], 
                               old_total_period, 
                               equal_sign=False), 
                'SUMIF($' + cells['CF_la_lag'] + ',">1",' + cells['dep_lag'] + ')'
                )  
        
        
        return cells_df_filled
        
                
    def __insert_rows__(self):
        
        ws = self.ws
        # num_years = self.new_num_years
        num_years = int(len(self.new_branch_df)/12)
        
        if num_years > 1:
            new_first_row = 20
            new_last_row = new_first_row + 12*(num_years - 1) - 1 
            
            Excel_Misc_Fns.format_range(
                ws, "row", new_first_row, new_last_row, 
                row_diff = -12, col_diff = 0)
        
        # save
        self.ws = ws
        
        
    def get_old_schedule_df(self):
        """
        Generate old schedule df from merging df summary & df formulae
        """
        
        old_lease_payment_df = self.get_old_df_summary().drop(columns = ['Contract']).copy()
        old_main_table_formulae_df = self.get_old_formulae_df().iloc[:,1:].copy()
        
        old_schedule_df = pd.concat([old_lease_payment_df,old_main_table_formulae_df],
                                    axis=1)
        
        self.old_schedule_df = old_schedule_df
        
        return old_schedule_df
    
    def get_new_schedule_df(self):
        """
        Generate new schedule df from merging df summary & df formulae 
        """
        
        new_lease_payment_df = self.get_new_df_summary().drop(columns = ['Contract']).copy()
        new_main_table_formulae_df = self.get_new_formulae_df().iloc[:,1:].copy()
        
        new_schedule_df = pd.concat([new_lease_payment_df,new_main_table_formulae_df],
                                    axis=1)
        
        self.new_schedule_df = new_schedule_df
        
        return new_schedule_df
    
        
                
    def write_meta_data(self):
        
        ws = self.ws
        df = self.df
        
        # branch name
        branch_name =self.contract
        
        # change sheet name to branch
        ws.title = branch_name[:30]
        
        # write the branch
        ws['A1'].value = branch_name
        
        # write the country 
        country = self.country
        ws['A2'].value = country 
        
        # save ws
        self.ws = ws
        
    def write_old_interest_rate(self):
        
        old_int_rate = self.df['Borrowing Rate (PFY)'].dropna().unique()
        
        if len(old_int_rate)==0:
            old_int_rate = np.nan
        else:
            old_int_rate = assert_one_and_get(old_int_rate)
            
        self.ws['S1'].value = old_int_rate
        
        self.old_int_rate = old_int_rate
        
    def write_new_interest_rate(self):
        
        df = self.df.copy()
        
        if self.type in ['Remeasurement','Addition']:
            new_int_rate = df['Borrowing Rate'].dropna().unique()[-1]
        else: 
            new_int_rate = self.old_int_rate
            
        self.ws['S2'].value = new_int_rate
    
    
    def write_old_period_formula(self):
        
        ws = self.ws
        first_row = self.first_row
        
        if self.type == 'Addition':
            formulae_df = (self.get_new_formulae_df()).loc[:,:0]
            formulae_df.loc[:,0] = ''
        else:
            formulae_df = (self.get_old_formulae_df()).loc[:,:0] 
            formulae_df.loc[0,0] = self.get_schedule_period_start()
        
        # write to excel
        Excel_Misc_Fns.df_to_worksheet(
            formulae_df, ws,
            index = False, header = False,
            startrow = first_row, startcol = 2)
        
    def write_new_period_formula(self):
        
        ws = self.ws
        first_row = self.first_row
        
        formulae_df = (self.get_new_formulae_df()).loc[:,:0] 
        
        if self.type == 'Remeasurement':
            formulae_df.loc[0,0] = self.get_schedule_period_start()
        
        # write to excel
        Excel_Misc_Fns.df_to_worksheet(
            formulae_df, ws,
            index = False, header = False,
            startrow = first_row, startcol = 11)
        
        
    def conditional_format(self):
        
        ws = self.ws
        
        try:
            row_no = self.remeasurement_row_idx
        except Exception:
            row_no = self.disposal_row_idx
        bg_format = PatternFill(start_color = 'C6E0B4', end_color = 'C6E0B4',
                                fill_type = 'solid')

        for row in ws.iter_rows(min_row=row_no, max_row=row_no, max_col=25):
            for cell in row:
                cell.fill = bg_format
            
    def write_total_lease_period(self):
        
        ws = self.ws
        
        ws['M2'].value = int(self.old_final_row_idx) - int(self.old_lease_start_row_idx) + 1
        
        if self.type=='Remeasurement':
            ws['N2'].value = int(self.new_final_row_idx) - int(self.remeasurement_row_idx) + 1
        else:
            ws['N2'].value = int(self.new_final_row_idx) - int(self.new_lease_start_row_idx) + 1
            
    def write_schedule_df(self):
        
        '''
        To write the full merged dataframe into excel 
        '''
        
        ws = self.ws
        first_row = self.first_row
        schedule_df = self.schedule_df
        
        cond = self.df['Type'].str.contains('Remeasurement')
        if cond.all():
            opening_rou = assert_one_and_get(self.df['Opening ROU (PFY)'].unique())
            opening_ll = assert_one_and_get(self.df['Opening Lease Liability (PFY)'].unique())
            schedule_df.loc[0,'W'] = ('={}'.format(opening_rou)) 
            schedule_df.loc[0,'S'] = ('={}'.format(opening_ll)) 
        
        Excel_Misc_Fns.df_to_worksheet(
            schedule_df, ws,
            index=False, header=False,
            startrow = first_row, startcol=1)
    
        

class OneContractDisclosure(OneContractSchedule):
    
    disc_amt_var_name = [
            'rou_start_pfy',
            'rou_addition_pfy',
            'rou_disposal_pfy',
            'll_rou_remeasurement_pfy',
            'rou_start',
            'rou_addition',
            'rou_disposal',
            'll_rou_remeasurement',
            'rou_cls',
            '',
            '',
            'acc_dep_p2fy',
            'dep_exp_pfy',
            'acc_dep_disp_pfy', 
            'acc_dep_pfy',
            'dep_exp',
            'acc_dep_disp',
            'acc_dep',
            '',
            '',
            'll_pfy_opn',
            'll_addtion_pfy',
            'll_disposal_pfy',
            'll_remeasurement_pfy',
            'interest_pfy',
            'payments_pfy',
            'neg_interest_pfy', 
            'll_opn',
            'll_addition',
            'll_disposal',
            'll_remeasurement',
            'interest',
            'payments',
            'neg_interest', 
            'll_cls',
            '',
            'll_curr',
            'll_non_curr',
            '',
            '',
            'payment_1',
            'payment_2',
            'payment_3',
            'payment_4',
            'payment_5',
            'payment_later',
            '',
            '',
            'interest_1',
            'interest_2',
            'interest_3',
            'interest_4',
            'interest_5',
            'interest_later',
            '',
            '',
            'npv_1',
            'npv_2',
            'npv_3',
            'npv_4',
            'npv_5',
            'npv_later'
            ]
     
    disc_date_var_name = [
            'start_date_1',
            'addition_date_pfy_1',
            'disposal_date_pfy_1',
            'remeasurement_date_pfy_1',
            'start_date_2',
            'addition_date_1',
            'disposal_date_1',
            'remeasurement_date_1',
            '',
            '',
            '',
            'start_date_3',
            'pfy_1',
            'disposal_date_pfy_2',
            'pfy_cls_1',
            'cfy_1',
            'disposal_date_2',
            'cfy_cls_1',
            '',
            '',
            'pfy_opn_1',
            'addition_date_pfy_2',
            'disposal_date_pfy_3',
            'remeasurement_date_pfy_2',
            'pfy_2',
            'pfy_3',
            'pfy_4',
            'cfy_opn',
            'addition_date_2',
            'disposal_date_3',
            'remeasurement_date_2',
            'cfy_2',
            'cfy_3',
            'cfy_4',
            'cfy_cls_2',
            '',
            'ffy_opn_to_cls',
            'ffy_cls0',
            '',
            '',
            'ffy_cls',
            'f2fy_cls',
            'f3fy_cls',
            'f4fy_cls',
            'f5fy_cls',
            'f6fy_cls',
            '',
            '',
            'ffy_cls2',
            'f2fy_cls2',
            'f3fy_cls2',
            'f4fy_cls2',
            'f5fy_cls2',
            'f6fy_cls2',
            '',
            '',
            'ffy_cls3',
            'f2fy_cls3',
            'f3fy_cls3',
            'f4fy_cls3',
            'f5fy_cls3',
            'f6fy_cls3',
            ]


    def __init__(self, df, contract, fy_start, fy_end, pfy_start, output_fp,):
                
        OneContractSchedule.__init__(self, df, contract, fy_start, fy_end, pfy_start, output_fp,)
        
        
    def __main__(self):
        
        print(f'Writing contract: {self.contract}...')
    
        OneContractSchedule.__main__(self)
        self.get_disclosure_date_formulae()
        self.get_disclosure_cell_formulae()        
        self.write_disclosure()
        self.format_disclosure()
        
        self.wb.save(self.fp)
        
        print(f"Completed.")

    
    
    def get_disclosure_cell_ref(self, col_header):
        """
        col_header: "Amount" / "Period" (column header for individual disclosure)
        """
        
        def fill_disclosure_dic(disc_var_name, start_row, start_col):
            """Stores dictionary of {variable_name: excel_cell_reference}"""
            disc_ref = {}
            row = start_row
            for var in disc_var_name:
                if var != "":
                    disc_ref[var] = Excel_Misc_Fns.xlref(row, start_col)
                else:
                    disc_ref['empty_'+str(row)] = Excel_Misc_Fns.xlref(row, start_col)
                row += 1
            return disc_ref
        
        if col_header == 'Amount':
            # 1. Generate dictionary of variable names and cell location.
            self.disc_amt_ref = fill_disclosure_dic(OneContractDisclosure.disc_amt_var_name, 10, 30) #AD10
            
            # 2. Create a dummy dataframe to store formulas later.
            self.disc_amt_ref_df = pd.Series(self.disc_amt_ref).to_frame(name = 'cell_ref')
            
        elif col_header == 'Period':
            # 1.
            self.disc_date_ref = fill_disclosure_dic(OneContractDisclosure.disc_date_var_name, 10, 32) #AF10
            
            # 2.
            self.disc_date_ref_df = pd.Series(self.disc_date_ref).to_frame(name = 'cell_ref')



    def get_disclosure_date_formulae(self): #add comments 
        """
        Stores a dictionary (self.disclosure_date_dict) of {variable_name: date}
        """
        # Load dummy dataframe
        self.get_disclosure_cell_ref('Period')
        

        dic = {}

        dic['start_date'] = pd.to_datetime(self.df['contract_start_date'].unique()[0]).date()
        dic['end_date'] = pd.to_datetime(self.df['contract_end_date'].unique()[0]).date()
        dic['addition_date'] = dic['start_date']
        dic['disposal_date'] = pd.to_datetime(self.df['disposal_date'].unique()[0]).date()
        dic['remeasurement_date'] = pd.to_datetime(self.df['remeasurement_date'].unique()[0]).date()
        
        dic['pfy_opn'] = self.fy_start - relativedelta(years=1)
        dic['pfy_cls'] = self.fy_end - relativedelta(years=1)
        dic['p2fy_cls'] = self.fy_end - relativedelta(years=2)
        dic['cfy_opn'] = self.fy_start
        dic['cfy_cls'] = self.fy_end
        dic['ffy_opn'] = self.fy_start + relativedelta(years=1)
        dic['ffy_cls'] = self.fy_end + relativedelta(years=1)
        dic['f2fy_cls'] = self.fy_end + relativedelta(years=2)
        dic['f3fy_cls'] = self.fy_end + relativedelta(years=3)
        dic['f4fy_cls'] = self.fy_end + relativedelta(years=4)
        dic['f5fy_cls'] = self.fy_end + relativedelta(years=5)
        dic['f6fy_cls'] = self.fy_end + relativedelta(years=6)
    
        self.disclosure_date_dict = dic.copy()


    def get_disclosure_cell_formulae(self):
        
        # Load dummy dataframe
        self.get_disclosure_cell_ref('Amount')
        
        # Rename variables for easier reference
        disc_ref = self.disc_amt_ref
        date_dict = self.disclosure_date_dict
        schedule_df = self.schedule_df.copy()
        sum_formula = Excel_Misc_Fns.sum_formula
        addition_formula = Excel_Misc_Fns.addition_formula
        subtraction_formula = Excel_Misc_Fns.subtraction_formula
        if_formula = Excel_Misc_Fns.if_formula
        isblank_formula = Excel_Misc_Fns.isblank_formula
        
        cond_remeasurement = self.df['Type'].str.contains('Remeasurement')
        cond_disposal = self.df['Type'].str.contains('Disposal')
        
        # Define variables that are frequently referenced 
        start_month = date_dict['start_date'].strftime('%b-%y')
        if not cond_remeasurement.any():
            start_month_idx = schedule_df.query(f"J == '{start_month}'").index[0] + self.first_row
        
        last_month_idx = schedule_df.query(f"J == 'Total'").index[0] - 1  + self.first_row
        
        cfy_start_str = date_dict['cfy_opn'].strftime("%b-%y")
        cfy_start_idx = schedule_df.query(f"J == '{cfy_start_str}'").index[0] + self.first_row
        
        cfy_end_str = date_dict['cfy_cls'].strftime("%b-%y")
        cfy_end_idx = schedule_df.query(f"J == '{cfy_end_str}'").index[0] + self.first_row
        pfy_start_idx = cfy_start_idx - 12
        pfy_end_idx = cfy_end_idx - 12
        
        
        # pfy_, cfy_
        # Based on input column alphabet and start and end index,
        # return string e.g. A1:A2. Only used with =SUM formula.
        
        def pfy_(col, start=pfy_start_idx, end=pfy_end_idx):
            return f"{col}{start}:{col}{end}"

        def cfy_(col, start=cfy_start_idx, end=cfy_end_idx):
            return f"{col}{start}:{col}{end}"
        
        ### Fill excel formulas for disclosure variables
        dic = {}
        
    # 1. Initial values
        # Contract starts before PFY open
        if date_dict['pfy_opn'] >= date_dict['start_date']:
        
            # Remeasurement table always starts from PFY open
            if cond_remeasurement.all():
                dic['rou_start_pfy'] = assert_one_and_get(self.df['Opening ROU (PFY)'].unique())
                dic['acc_dep_p2fy'] = assert_one_and_get(self.df['Opening Accumulated Depreciation (PFY)'].unique())
                dic['ll_pfy_opn'] = assert_one_and_get(self.df['Opening Lease Liability (PFY)'].unique())
            
            else:
                dic['rou_start_pfy'] = f"=W{start_month_idx}" # f"=INDEX(W{self.first_row}:W10000,MATCH(TRUE,INDEX((W{self.first_row}:W10000<>0),0),0))"
                dic['acc_dep_p2fy'] = sum_formula(f"X{start_month_idx}", f"X{pfy_start_idx-1}") + "*-1" # f"=-SUM(X{self.first_row}:X{pfy_start_idx-1})"
                dic['ll_pfy_opn'] = f"=S{pfy_start_idx}"
        
        # Contract starts after PFY open -> 0
        else:
            dic['rou_start_pfy'] = 0
            dic['acc_dep_p2fy'] = 0
            dic['ll_pfy_opn'] = 0
    
    # 2. Common formulas regardless of contract types
        dic['ll_remeasurement_pfy'] = self.df['Remeasurements'].fillna(0).max() # Value should be provided in the input template.
            
        dic['rou_start'] = sum_formula(disc_ref['rou_start_pfy'], disc_ref['ll_rou_remeasurement_pfy'])
        dic['rou_cls'] = sum_formula(disc_ref['rou_start'], disc_ref['ll_rou_remeasurement'])
        
        dic['acc_dep_pfy'] = sum_formula(disc_ref['acc_dep_p2fy'], disc_ref['acc_dep_disp_pfy'])
        dic['acc_dep'] = sum_formula(disc_ref['acc_dep_pfy'], disc_ref['acc_dep_disp'])
        
        dic['neg_interest_pfy'] = "=" + disc_ref['interest_pfy'] + "* -1"
        dic['neg_interest'] = "=" + disc_ref['interest'] + "* -1"
        
        dic['ll_opn'] = sum_formula(disc_ref['ll_pfy_opn'], disc_ref['neg_interest_pfy'])
        dic['ll_cls'] = sum_formula(disc_ref['ll_opn'], disc_ref['neg_interest'])
        
        dic['dep_exp'] = f"=SUM({cfy_('X')})" + "* -1"
        dic['interest'] = f"=SUM({cfy_('U')})"
        dic['payments'] = subtraction_formula([f"SUM({cfy_('T')})", disc_ref['neg_interest']])
    
        dic['dep_exp_pfy'] = f"=SUM({pfy_('X')})" + "* -1" 
        dic['interest_pfy'] = f"=SUM({pfy_('U')})"
        dic['payments_pfy'] = subtraction_formula([f"SUM({pfy_('T')})", disc_ref['neg_interest_pfy']])
        
    # 3A. Addition
        # Contract starts in PFY (after open, before close)
        if date_dict['start_date'] >= date_dict['pfy_opn'] and date_dict['start_date'] <= date_dict['pfy_cls']:
            dic['rou_addition_pfy'] = f"=W{start_month_idx}" # f"=INDEX({pfy_('W')}, MATCH(TRUE, INDEX({pfy_('W')}<>0,),0))"
            dic['ll_addtion_pfy'] = f"=S{start_month_idx}" # f"=INDEX({pfy_('S')}, MATCH(TRUE, INDEX({pfy_('S')}<>0,),0))"
            dic['rou_addition'] = 0
            dic['ll_addition'] = 0
                
        # Contract starts in CFY (after open, before close)
        elif date_dict['start_date'] >= date_dict['cfy_opn'] and date_dict['start_date'] <= date_dict['cfy_cls']:
            dic['rou_addition_pfy'] = 0
            dic['ll_addtion_pfy'] = 0
            dic['rou_addition'] = f"=W{start_month_idx}" # f"=INDEX({cfy_('W')}, MATCH(TRUE, INDEX({cfy_('W')}<>0,),0))"
            dic['ll_addition'] = f"=S{start_month_idx}" # f"=INDEX({cfy_('S')}, MATCH(TRUE, INDEX({cfy_('S')}<>0,),0))"
            
            dic['dep_exp_pfy'] = 0
            dic['interest_pfy'] = 0
            dic['payments_pfy'] = 0
            
        else:
            dic['rou_addition_pfy'] = 0
            dic['ll_addtion_pfy'] = 0
            dic['rou_addition'] = 0
            dic['ll_addition'] = 0
        
        
            
    # 3B. Disposal
        
        if cond_disposal.any():
            disp_date = date_dict['disposal_date'].strftime("%b-%y")
            disp_date_idx = schedule_df.query(f"J =='{disp_date}'").index[0] + self.first_row
            
            if date_dict['disposal_date'] < date_dict['pfy_cls']: # if contract disposed in PFY
                dic['rou_disposal'] = 0
                dic['acc_dep_disp'] = 0
                dic['ll_disposal'] = 0
                dic['rou_disposal_pfy'] = \
                        if_formula(
                            isblank_formula(disc_ref['rou_start_pfy'], equal_sign=False), 
                            disc_ref['rou_addition_pfy'], 
                            disc_ref['rou_start_pfy']) + "*-1"
                dic['acc_dep_disp_pfy'] = "=-1*" + addition_formula([disc_ref['acc_dep_p2fy'], disc_ref['dep_exp_pfy']], equal_sign=False)
                dic['ll_disposal_pfy'] = f"=-V{disp_date_idx}" #f"=INDEX(V:V, MATCH('{disp_date}', A:A, 0),)"
                
                dic['dep_exp'] = 0
                dic['interest'] = 0
                dic['payments'] = 0
                
                dic['dep_exp_pfy'] = sum_formula(f"X{pfy_start_idx}", f"X{disp_date_idx}") + "* -1"
                dic['interest_pfy'] = sum_formula(f"U{pfy_start_idx}", f"U{disp_date_idx}")
                dic['payments_pfy'] = subtraction_formula([sum_formula(f"T{pfy_start_idx}", f"T{disp_date_idx}", equal_sign=False), 
                                                           disc_ref['neg_interest_pfy']])
                
            elif date_dict['disposal_date'] < date_dict['cfy_cls']: # if contract disposed in CFY
                dic['rou_disposal'] = \
                        if_formula(
                            isblank_formula(disc_ref['rou_start'], equal_sign=False), 
                            disc_ref['rou_addition'], 
                            disc_ref['rou_start']) + "*-1"
                dic['acc_dep_disp'] = "=-1*" + addition_formula([disc_ref['acc_dep_pfy'], disc_ref['dep_exp']], equal_sign=False)
                dic['ll_disposal'] = f"=-V{disp_date_idx}" #f"=INDEX(V:V, MATCH('{disp_date}', A:A, 0),)"
                dic['rou_disposal_pfy'] = 0
                dic['acc_dep_disp_pfy'] = 0
                dic['ll_disposal_pfy'] = 0
                
                dic['dep_exp'] = sum_formula(f"X{cfy_start_idx}", f"X{disp_date_idx}") + "* -1"
                dic['interest'] = sum_formula(f"U{cfy_start_idx}", f"U{disp_date_idx}")
                dic['payments'] = subtraction_formula([sum_formula(f"T{cfy_start_idx}", f"T{disp_date_idx}", equal_sign=False),
                                                       disc_ref['neg_interest']])
                
        else:
            dic['rou_disposal'] = 0
            dic['acc_dep_disp'] = 0
            dic['ll_disposal'] = 0
            dic['rou_disposal_pfy'] = 0
            dic['acc_dep_disp_pfy'] = 0
            dic['ll_disposal_pfy'] = 0
            
            
            
    # 3C. Remeasurement
        if cond_remeasurement.any():
            rm_date = date_dict['remeasurement_date'].strftime("%b-%y")
            date_index = schedule_df.query(f"J == '{rm_date}'").index[0] + self.first_row # f'MATCH("{date_}", A:A, 0)'
            sum_npv_new = sum_formula(f"R{date_index}", f"R{last_month_idx}", equal_sign=False) # f'SUM(INDIRECT(CONCATENATE("R", {date_index})):INDIRECT(CONCATENATE("R", MATCH("Total", A:A, 0)-1)))'
            dic['ll_remeasurement'] = subtraction_formula([sum_npv_new, f"S{date_index}"])
        else:
            dic['ll_remeasurement'] = 0
            
            

    # 4. Summary calculations at the bottom
        nfy_start = (self.fy_start + relativedelta(years=1)).strftime("%b-%y")
        if len(schedule_df.query(f"A == '{nfy_start}'"))>0:
            nfy_start_idx = schedule_df.query(f"A == '{nfy_start}'").index[0] + self.first_row
        elif len(schedule_df.query(f"A == 'Total'"))>0:
            nfy_start_idx = schedule_df.query(f"A == 'Total'").index[0] + self.first_row + 1
        else:
            nfy_start_idx = schedule_df.query(f"J == 'Total'").index[0] + self.first_row + 1
            
        if (date_dict['end_date'] <= date_dict['cfy_cls'] or 
            self.type == 'Disposal'):
            dic['ll_curr'] = 0
            dic['ll_non_curr'] = 0
        else:
            dic['ll_curr'] = f'=S{nfy_start_idx} - V{nfy_start_idx+11}'
            dic['ll_non_curr'] = f'=V{nfy_start_idx+11}'
            
    
        fy_end_ = self.fy_end
        end_date_ = self.disclosure_date_dict['end_date'].strftime("%b-%y")
        contract_end_idx = schedule_df.query(f"J == '{end_date_}'").index[0] + self.first_row # "MATCH({end_date_}, A:A, 0)"
        
        ## Following code is used to populate 'payments_1/2/3/4/5/later' and 'interest_1/2/3/4/5/later'
        i = 1
        while i < 7:
            
            contract_ended = (fy_end_ + relativedelta(years=i)).year > self.disclosure_date_dict['end_date'].year
             
            if (not contract_ended and i < 6):
            # if contract ends after 'current' year:
                start_ = nfy_start_idx + (i-1)*12
                end_ = start_ + 11
            elif (not contract_ended and i == 6):
                start_ = nfy_start_idx + (i-1)*12
                end_ = contract_end_idx
            else:
            # if contract ends in the 'current' year:
                start_ = nfy_start_idx + (i-1)*12
                end_ = contract_end_idx
                
            if i == 6:
                dic['payment_later'] = sum_formula(f"T{start_}", f"T{end_}", negative=True) # f'=-SUM(T{start_}:T{end_})'
                dic['interest_later'] = sum_formula(f"U{start_}", f"U{end_}", negative=True) # f'=-SUM(U{start_}:U{end_})'
            else:
                dic['payment_'+str(i)] = sum_formula(f"T{start_}", f"T{end_}", negative=True) # f'=-SUM(T{start_}:T{end_})'
                dic['interest_'+str(i)] = sum_formula(f"U{start_}", f"U{end_}", negative=True) # f'=-SUM(U{start_}:U{end_})'
            
            if contract_ended or self.type=='Disposal':
                if cond_disposal.any():
                    i = 1
                while i < 7:
                    if i == 6:
                        dic['payment_later'] = 0
                        dic['interest_later'] = 0
                        break
                    else:
                        dic['payment_'+str(i)] = 0
                        dic['interest_'+str(i)] = 0
                        i += 1
                break
            else:
                i += 1
            
        dic['npv_1'] = addition_formula([disc_ref['payment_1'], disc_ref['interest_1']])
        dic['npv_2'] = addition_formula([disc_ref['payment_2'], disc_ref['interest_2']])
        dic['npv_3'] = addition_formula([disc_ref['payment_3'], disc_ref['interest_3']])
        dic['npv_4'] = addition_formula([disc_ref['payment_4'], disc_ref['interest_4']])
        dic['npv_5'] = addition_formula([disc_ref['payment_5'], disc_ref['interest_5']])
        dic['npv_later'] = addition_formula([disc_ref['payment_later'], disc_ref['interest_later']])
        
        self.disclosure_formulae_dict = dic.copy()
        
        # END

    def load_amt_formulae_to_df(self):
        
        disc_ref = self.disc_amt_ref
        date_dict = self.disclosure_date_dict
        schedule_df = self.schedule_df.copy()
        sum_formula = Excel_Misc_Fns.sum_formula
        addition_formula = Excel_Misc_Fns.addition_formula
        subtraction_formula = Excel_Misc_Fns.subtraction_formula
        if_formula = Excel_Misc_Fns.if_formula
        isblank_formula = Excel_Misc_Fns.isblank_formula
        
        disc_df = self.disc_amt_ref_df.copy()
        formula_ = self.disclosure_formulae_dict.copy()
        
        disc_df['cell_value'] = np.nan
        disc_df['cell_value'] = disc_df['cell_value'].astype(object)
    
        disc_df.at['rou_start_pfy', 'cell_value'] = formula_['rou_start_pfy']
        disc_df.at['rou_addition_pfy', 'cell_value'] = formula_['rou_addition_pfy']
        disc_df.at['rou_disposal_pfy', 'cell_value'] = formula_['rou_disposal_pfy']
        disc_df.at['ll_rou_remeasurement_pfy', 'cell_value'] = formula_['ll_remeasurement_pfy']
        disc_df.at['rou_start', 'cell_value'] = formula_['rou_start']
        disc_df.at['rou_addition', 'cell_value'] = formula_['rou_addition']
        disc_df.at['rou_disposal', 'cell_value'] = formula_['rou_disposal']
        disc_df.at['ll_rou_remeasurement', 'cell_value'] = formula_['ll_remeasurement']
        disc_df.at['rou_cls', 'cell_value'] = formula_['rou_cls']
        disc_df.at['acc_dep_p2fy', 'cell_value'] = formula_['acc_dep_p2fy']
        disc_df.at['dep_exp_pfy', 'cell_value'] = formula_['dep_exp_pfy']
        disc_df.at['acc_dep_disp_pfy', 'cell_value'] = formula_['acc_dep_disp_pfy']
        disc_df.at['acc_dep_pfy', 'cell_value'] = formula_['acc_dep_pfy']
        disc_df.at['dep_exp', 'cell_value'] = formula_['dep_exp']
        disc_df.at['acc_dep_disp', 'cell_value'] = formula_['acc_dep_disp']
        disc_df.at['acc_dep', 'cell_value'] = formula_['acc_dep']
        disc_df.at['ll_pfy_opn', 'cell_value'] = formula_['ll_pfy_opn']
        disc_df.at['ll_addtion_pfy', 'cell_value'] = formula_['ll_addtion_pfy']
        disc_df.at['ll_disposal_pfy', 'cell_value'] = formula_['ll_disposal_pfy']
        disc_df.at['ll_remeasurement_pfy', 'cell_value'] = formula_['ll_remeasurement_pfy']
        disc_df.at['interest_pfy', 'cell_value'] = formula_['interest_pfy']
        disc_df.at['payments_pfy', 'cell_value'] = formula_['payments_pfy']
        disc_df.at['neg_interest_pfy', 'cell_value'] = formula_['neg_interest_pfy']
        disc_df.at['ll_opn', 'cell_value'] = formula_['ll_opn']
        disc_df.at['ll_addition', 'cell_value'] = formula_['ll_addition']
        disc_df.at['ll_disposal', 'cell_value'] = formula_['ll_disposal']
        disc_df.at['ll_remeasurement', 'cell_value'] = formula_['ll_remeasurement']
        disc_df.at['interest', 'cell_value'] = formula_['interest']
        disc_df.at['payments', 'cell_value'] = formula_['payments']
        disc_df.at['neg_interest', 'cell_value'] = formula_['neg_interest']
        disc_df.at['ll_cls', 'cell_value'] = formula_['ll_cls']
        disc_df.at['ll_curr', 'cell_value'] = formula_['ll_curr']
        disc_df.at['ll_non_curr', 'cell_value'] = formula_['ll_non_curr']
        disc_df.at['payment_1', 'cell_value'] = formula_['payment_1']
        disc_df.at['payment_2', 'cell_value'] = formula_['payment_2']
        disc_df.at['payment_3', 'cell_value'] = formula_['payment_3']
        disc_df.at['payment_4', 'cell_value'] = formula_['payment_4']
        disc_df.at['payment_5', 'cell_value'] = formula_['payment_5']
        disc_df.at['payment_later', 'cell_value'] = formula_['payment_later']
        disc_df.at['interest_1', 'cell_value'] = formula_['interest_1']
        disc_df.at['interest_2', 'cell_value'] = formula_['interest_2']
        disc_df.at['interest_3', 'cell_value'] = formula_['interest_3']
        disc_df.at['interest_4', 'cell_value'] = formula_['interest_4']
        disc_df.at['interest_5', 'cell_value'] = formula_['interest_5']
        disc_df.at['interest_later', 'cell_value'] = formula_['interest_later']
        disc_df.at['npv_1', 'cell_value'] = formula_['npv_1']
        disc_df.at['npv_2', 'cell_value'] = formula_['npv_2']
        disc_df.at['npv_3', 'cell_value'] = formula_['npv_3']
        disc_df.at['npv_4', 'cell_value'] = formula_['npv_4']
        disc_df.at['npv_5', 'cell_value'] = formula_['npv_5']
        disc_df.at['npv_later', 'cell_value'] = formula_['npv_later']
        
        return disc_df

    
    def load_date_formulae_to_df(self):
        
        disc_ref = self.disc_amt_ref
        date_dict = self.disclosure_date_dict
        schedule_df = self.schedule_df.copy()
        sum_formula = Excel_Misc_Fns.sum_formula
        addition_formula = Excel_Misc_Fns.addition_formula
        subtraction_formula = Excel_Misc_Fns.subtraction_formula
        if_formula = Excel_Misc_Fns.if_formula
        isblank_formula = Excel_Misc_Fns.isblank_formula
        
        date_df = self.disc_date_ref_df.copy()
        date_df_2 = date_df.copy()
        
        date_dict = self.disclosure_date_dict.copy()
        cond_remeasurement = self.df['Type'].str.contains('Remeasurement')
        cond_disposal = self.df['Type'].str.contains('Disposal')
        
        # Contract starts before CFY
        if date_dict['start_date'] < date_dict['cfy_opn']:
            
            # Contract starts before PFY open
            if date_dict['start_date'] < date_dict['pfy_opn']:
                date_df.at['start_date_1', 'cell_value'] = date_dict['start_date']
                date_df.at['start_date_3', 'cell_value'] = date_dict['start_date']
                
            # Contract starts in PFY
            else:
                date_df.at['addition_date_pfy_1', 'cell_value'] = date_dict['addition_date']
                date_df.at['addition_date_pfy_2', 'cell_value'] = date_dict['addition_date']
            
            date_df.at['pfy_1', 'cell_value'] = date_dict['pfy_opn']
            date_df_2.at['pfy_1', 'cell_value'] = date_dict['pfy_cls']
            date_df.at['pfy_cls_1', 'cell_value'] = date_dict['cfy_opn']
            
            date_df.at['pfy_opn_1', 'cell_value'] = date_dict['pfy_opn']
            date_df.at['pfy_2', 'cell_value'] = date_dict['pfy_opn']
            date_df.at['pfy_3', 'cell_value'] = date_dict['pfy_opn']
            date_df.at['pfy_4', 'cell_value'] = date_dict['pfy_opn']
            date_df_2.at['pfy_2', 'cell_value'] = date_dict['pfy_cls']
            date_df_2.at['pfy_3', 'cell_value'] = date_dict['pfy_cls']
            date_df_2.at['pfy_4', 'cell_value'] = date_dict['pfy_cls']
        
        # Contract starts in CFY
        elif date_dict['start_date'] < date_dict['cfy_cls']:
            date_df.at['addition_date_1', 'cell_value'] = date_dict['addition_date']
            date_df.at['addition_date_2', 'cell_value'] = date_dict['addition_date']
        
        
        if cond_disposal.any():
            # Contract disposed in PFY
            if date_dict['disposal_date'] < date_dict['cfy_opn']:
                date_df.at['disposal_date_pfy_1', 'cell_value'] = date_dict['disposal_date']
                date_df.at['disposal_date_pfy_2', 'cell_value'] = date_dict['disposal_date']
                date_df.at['disposal_date_pfy_3', 'cell_value'] = date_dict['disposal_date']
            
            # Contract disposed in CFY
            elif date_dict['disposal_date'] < date_dict['cfy_cls']:
                date_df.at['disposal_date_1', 'cell_value'] = date_dict['disposal_date']
                date_df.at['disposal_date_2', 'cell_value'] = date_dict['disposal_date']
                date_df.at['disposal_date_3', 'cell_value'] = date_dict['disposal_date']


        if cond_remeasurement.any():
            date_df.at['remeasurement_date_1', 'cell_value'] = date_dict['remeasurement_date']
            date_df.at['remeasurement_date_2', 'cell_value'] = date_dict['remeasurement_date']
        
        # if date_dict['start_date'] < date_dict['cfy_opn'] and date_dict['start_date'] > date_dict['pfy_opn']: 
        #     date_df.at['start_date_2', 'cell_value'] = if_formula("AND(ISBLANK(AZ10), ISBLANK(AZ11))", date_dict['start_date'], )
    
        if (not cond_disposal.any()) or (date_dict['disposal_date'] > date_dict['cfy_opn']):
        # Contract is not Disposed at all, or is Disposed in CFY
            date_df.at['cfy_1', 'cell_value'] = date_dict['cfy_opn']
            date_df_2.at['cfy_1', 'cell_value'] = date_dict['cfy_cls']
            date_df.at['cfy_cls_1', 'cell_value'] = date_dict['cfy_cls']
        
            date_df.at['cfy_opn', 'cell_value'] = date_dict['cfy_opn']
            date_df.at['cfy_2', 'cell_value'] = date_dict['cfy_opn']
            date_df.at['cfy_3', 'cell_value'] = date_dict['cfy_opn']
            date_df.at['cfy_4', 'cell_value'] = date_dict['cfy_opn']
            date_df_2.at['cfy_2', 'cell_value'] = date_dict['cfy_cls']
            date_df_2.at['cfy_3', 'cell_value'] = date_dict['cfy_cls']
            date_df_2.at['cfy_4', 'cell_value'] = date_dict['cfy_cls']
            date_df.at['cfy_cls_2', 'cell_value'] = date_dict['cfy_cls']
        
            date_df.at['ffy_opn_to_cls', 'cell_value'] = date_dict['ffy_opn']
            date_df_2.at['ffy_opn_to_cls', 'cell_value'] = date_dict['ffy_cls']
            date_df.at['ffy_cls0', 'cell_value'] = date_dict['ffy_cls']
        
            ## Following code is used to populate 'fxfy_cls' non-manually
            i = 1
            while i < 7:
                if cond_disposal.any():
                    break
                if i == 1:
                    i = ""
                if date_dict[f"f{i}fy_cls"].year <= date_dict['end_date'].year:
                    date_df.at[f"f{i}fy_cls", 'cell_value'] = date_dict[f"f{i}fy_cls"]
                    date_df.at[f"f{i}fy_cls2", 'cell_value'] = date_dict[f"f{i}fy_cls"]
                    date_df.at[f"f{i}fy_cls3", 'cell_value'] = date_dict[f"f{i}fy_cls"]
                    if i == "":
                        i = 1
                else:
                    break
                i += 1
        
        # Combine date_df and date_df_2 to get Period and Period 2 columns respectively
        date_df = date_df.merge(date_df_2, how = 'left', on = 'cell_ref')
        date_df['cell_value_x'] = pd.to_datetime(date_df['cell_value_x']).dt.strftime("%b-%y") # Period
        date_df['cell_value_y'] = pd.to_datetime(date_df['cell_value_y']).dt.strftime("%b-%y") # Period 2
        
        return date_df
    
    
    def load_headers_to_df(self):
        
        header_df = pd.read_excel(self.fp, sheet_name='branch_disclosure_template', header=None)
        
        return header_df
          
    
    def write_disclosure(self):
        
        header_df = self.load_headers_to_df()
        self.date_df = self.load_date_formulae_to_df()
        self.disc_df = self.load_amt_formulae_to_df()
        
        # Output the dataframes to worksheet
        Excel_Misc_Fns.df_to_worksheet(
            header_df, self.ws, 
            header=False, index=False, 
            startrow = 9, startcol=28) #AB9
        
        Excel_Misc_Fns.df_to_worksheet(
            self.disc_df[['cell_value']], self.ws, 
            header=False, index=False, 
            startrow = 10, startcol=30) #AD10
        
        Excel_Misc_Fns.df_to_worksheet(
            self.date_df[['cell_value_x', 'cell_value_y']], self.ws, 
            header=False, index=False, 
            startrow = 10, startcol=32) #AF10
        
    def format_disclosure(self):
        
        # 1. Fill <PREV FY> and <FY> values.
        row_, col_ = 9,'AB' #AB9 - "Names"
        
        none_count = 0
        for r in self.ws[col_][row_-1:]:
            if r.value is None or r.value is np.nan:
                none_count += 1
            else:
                if "<PREV FY START>" in r.value:
                    r.value = r.value.replace("<PREV FY START>", 
                                              (self.fy_start - relativedelta(years=1)).strftime("%d %B %Y"))
                elif "<PREV FY END>" in r.value:
                    r.value = r.value.replace("<PREV FY END>", 
                                              (self.fy_end - relativedelta(years=1)).strftime("%d %B %Y"))
                elif "<FY END>" in r.value:
                    r.value = r.value.replace("<FY END>", 
                                              self.fy_end.strftime("%d %B %Y"))
            if none_count > 4:
                break
                
        # 2. Hide columns
        self.ws.column_dimensions['AC'].hidden = True
        self.ws.column_dimensions['AE'].hidden = True
        
        # # 3. Fix column format
        # self.ws.column_dimensions['AD'].number_format = BUILTIN_FORMATS[39] # '#,##0.00_);(#,##0.00)'
        
        for c in ['AB9', 'AD9', 'AF9', 'AG9']:
            self.ws[c].font = Font(name='Arial', size=10, bold=True, underline='single')
        #     self.ws.column_dimensions[col].font = Font(name='Arial', size=10, bold=False, italic=False, underline='none')
        #     self.ws.column_dimensions[col].fill = PatternFill()

    
class AllDisclosure:
    
    VARIABLE_NAME = [
        'blank',
        'rou_start_pfy',
        'rou_addition_pfy',
        'rou_disposal_pfy',
        'll_rou_remeasurement_pfy',
        'rou_others_pfy', #blank',
        'rou_forex_pfy',
        'rou_start',
        'rou_addition',
        'rou_disposal',
        'll_rou_remeasurement',
        'rou_others', #'blank',
        'rou_forex',
        'rou_cls',
        'blank',
        'blank',
        'acc_dep_start_pfy',
        'dep_exp_pfy',
        'acc_dep_disp_pfy',
        'acc_dep_forex',
        'acc_dep_start',
        'dep_exp',
        'acc_dep_disp',
        'acc_dep_impairment', #'blank',
        'acc_dep_others', #'blank',
        'acc_dep_forex',
        'acc_dep',
        'blank',
        'blank',
        'net_rou_start_pfy',
        'net_rou_start',
        'net_rou',
        'blank',
        'blank',
        'll_start_pfy',
        'll_addtion_pfy',
        'll_disposal_pfy',
        'll_rou_remeasurement_pfy',
        'interest_pfy',
        'payments_pfy',
        'blank',
        'neg_interest_pfy',
        'll_forex_pfy',
        'll_start',
        'll_addition',
        'll_disposal',
        'll_rou_remeasurement',
        'interest',
        'payments',
        'rental_rebate', #'blank',
        'neg_interest',
        'll_forex',
        'll_cls',
        'blank',
        'll_curr',
        'll_non_curr',
        'll_total',
        'check_1',
        'blank',
        'blank',
        'blank',
        'blank',
        'payment_1',
        'payment_2',
        'payment_3',
        'payment_4',
        'payment_5',
        'payment_later',
        'payment_total',
        'blank',
        'blank',
        'interest_1',
        'interest_2',
        'interest_3',
        'interest_4',
        'interest_5',
        'interest_later',
        'interest_total',
        'blank',
        'blank',
        'npv_1',
        'npv_2',
        'npv_3',
        'npv_4',
        'npv_5',
        'npv_later',
        'npv_total',
        'check_2',
        'blank',
        'blank',
        'blank',
        'blank',
        'blank',
        'blank',
        'blank',
        'blank',
        'blank',
        'blank',
        'blank',
        'blank',
        'blank',
        'blank',
        'blank',
        'blank',
        'blank',
        'forex_start_pfy',
        'forex_avg_pfy',
        'fores_start',
        'forex_avg',
        'forex_cls']
        
    def get_formulae(self):
            
        curr_col = self.curr_col
        input_df = self.contract_value['input_df']
        
        # Not constant, hence using a method to generate a variable formulaes
        # based on contract code
        
        DISC_FORMULA = [
            # ROU
            np.nan,
            f"= '{self.contract}'!AD10*{curr_col}113", 
            f"= '{self.contract}'!AD11*{curr_col}114",
            f"= '{self.contract}'!AD12*{curr_col}114",
            f"= '{self.contract}'!AD13*{curr_col}114",
            0,
            f'= {curr_col}15-SUM({curr_col}9:{curr_col}13)',
            f"= '{self.contract}'!AD14*{curr_col}115", # f"= IF({curr_col}11=0,'{self.contract}'!V14,0)*{curr_col}115", %
            f"= '{self.contract}'!AD15*{curr_col}116",
            f"= '{self.contract}'!AD16*{curr_col}116",
            f"= '{self.contract}'!AD17*{curr_col}116",
            0,
            f'= {curr_col}21-SUM({curr_col}15:{curr_col}19)',
            f"= '{self.contract}'!AD18*{curr_col}117",# f"= IF(C17=0,('{self.contract}'!V14+'{self.contract}'!AG18),0)*C117", % 
            np.nan,
            np.nan,
            
            # AccDep
            f"= '{self.contract}'!AD21*{curr_col}113", 
            f"= '{self.contract}'!AD22*{curr_col}114", # f"= -SUM('{self.contract}'!W14:W19)*C114",
            f"= '{self.contract}'!AD23*{curr_col}114",
            f'= {curr_col}28-SUM({curr_col}24:{curr_col}26)',
            f"= '{self.contract}'!AD24*{curr_col}115", # f"= IF(C26=0,-SUM('{self.contract}'!W14:W19),0)*C115", %
            f"= '{self.contract}'!AD25*{curr_col}116", # f"= -SUM('{self.contract}'!W20:W31)*C116",
            f"= '{self.contract}'!AD26*{curr_col}116",
            np.nan,
            np.nan,
            f'= {curr_col}34-SUM({curr_col}28:{curr_col}32)', 
            f"= '{self.contract}'!AD27*{curr_col}117",# f"= IF(C30=0,-SUM('{self.contract}'!W14:W31),0)*C117", % 
            np.nan,
            np.nan,
            
            # Carrying value
            f'= {curr_col}9-{curr_col}24', 
            f'= {curr_col}15-{curr_col}28',
            f'= {curr_col}21-{curr_col}34',
            np.nan,
            np.nan,
            
            # Lease Liabilities
            f"= '{self.contract}'!AD30*{curr_col}113", 
            f"= '{self.contract}'!AD31*{curr_col}114",
            f"= '{self.contract}'!AD32*{curr_col}114",
            f"= '{self.contract}'!AD33*{curr_col}114",
            f"= '{self.contract}'!AD34*{curr_col}114",
            f"= '{self.contract}'!AD35*{curr_col}114 - {curr_col}48", # f"= SUM('{self.contract}'!S14:S19)*C114-SUM(C48:C49)", % 
            np.nan,
            f"= '{self.contract}'!AD36*{curr_col}114",
            f'= {curr_col}51-SUM({curr_col}42:{curr_col}49)',
            f"= '{self.contract}'!AD37*{curr_col}115", # f"= IF(C44=0,'{self.contract}'!U19,0)*C115",
            f"= '{self.contract}'!AD38*{curr_col}116",
            f"= '{self.contract}'!AD39*{curr_col}116",
            f"= '{self.contract}'!AD40*{curr_col}116",
            f"= '{self.contract}'!AD41*{curr_col}116",
            f"= '{self.contract}'!AD42*{curr_col}116 - {curr_col}57", # f"= SUM('{self.contract}'!S20:S31)*C116-SUM(C57:C58)", %
            np.nan,
            f"= '{self.contract}'!AD43*{curr_col}116",
            f'= {curr_col}60-SUM({curr_col}51:{curr_col}58)',
            f"= '{self.contract}'!AD44*{curr_col}117", #f"= IF(C53=0,'{self.contract}'!U31,0)*C117", % 
            np.nan,
            f"= '{self.contract}'!AD46*{curr_col}117", # f"= ('{self.contract}'!R32-'{self.contract}'!U37)*C117",
            f"= '{self.contract}'!AD47*{curr_col}117",
            f'= SUM({curr_col}62:{curr_col}63)',
            f'= ROUND({curr_col}60-{curr_col}64,0)',
            np.nan,
            np.nan,
            np.nan,
            np.nan,
            f"= '{self.contract}'!AD50*{curr_col}117",
            f"= '{self.contract}'!AD51*{curr_col}117",
            f"= '{self.contract}'!AD52*{curr_col}117",
            f"= '{self.contract}'!AD53*{curr_col}117",
            f"= '{self.contract}'!AD54*{curr_col}117",
            f"= '{self.contract}'!AD55*{curr_col}117",
            f'= SUM({curr_col}70:{curr_col}75)',
            np.nan,
            np.nan,
            f"= '{self.contract}'!AD58*{curr_col}117",
            f"= '{self.contract}'!AD59*{curr_col}117",
            f"= '{self.contract}'!AD60*{curr_col}117",
            f"= '{self.contract}'!AD61*{curr_col}117",
            f"= '{self.contract}'!AD62*{curr_col}117",
            f"= '{self.contract}'!AD63*{curr_col}117",
            f'= SUM({curr_col}79:{curr_col}84)',
            np.nan,
            np.nan,
            f"= '{self.contract}'!AD66*{curr_col}117",
            f"= '{self.contract}'!AD67*{curr_col}117",
            f"= '{self.contract}'!AD68*{curr_col}117",
            f"= '{self.contract}'!AD69*{curr_col}117",
            f"= '{self.contract}'!AD70*{curr_col}117",
            f"= '{self.contract}'!AD71*{curr_col}117",
            f'= SUM({curr_col}88:{curr_col}93)',
            f'= ROUND({curr_col}60-{curr_col}94,0)',
            np.nan,
            np.nan,
            np.nan,
            np.nan,
            np.nan,
            np.nan,
            np.nan,
            f'= {curr_col}104-SUM({curr_col}98:{curr_col}102)',
            np.nan,
            np.nan,
            np.nan,
            np.nan,
            np.nan,
            f'= {curr_col}110-SUM({curr_col}104:{curr_col}108)',
            np.nan,
            np.nan,
            np.nan,
            assert_one_and_get(input_df['Opening ER (PFY)'].unique()), # f"= '{self.contract}'!H15",
            assert_one_and_get(input_df['Average ER (PFY)'].unique()), # f"= '{self.contract}'!I15",
            assert_one_and_get(input_df['Opening ER'].unique()), # f"= '{self.contract}'!J15",
            assert_one_and_get(input_df['Average ER'].unique()), # f"= '{self.contract}'!K15",
            assert_one_and_get(input_df['Closing ER'].unique()) # f"= '{self.contract}'!L15"
            ]
        
        FORMULA_DF = pd.DataFrame({self.contract: DISC_FORMULA})
        
        return FORMULA_DF
    
    def __init__(self, client, contract_dict, output_fp, pfy_start, pfy_end, fy_end):
        
        self.client = client
        self.pfy_start = pfy_start
        self.pfy_end = pfy_end
        self.fy_end = fy_end
        self.output_fp = output_fp
        self.contracts = contract_dict
        
    def __main__(self):
        
        self.wb = openpyxl.load_workbook(self.output_fp)
        self.ws_lead = self.wb.copy_worksheet(self.wb['lead_template'])
        self.ws_lead.title = "lead sheet"
        
        self.total_col = 3 # Column C
        self.curr_col_num = 2 # Column B
        self.load_all_disc()
        self.format_all_disc()
        
        self.wb.save(self.output_fp)
        self.wb.close()
        
        print(f"Disclosure for all contracts completed.")
        
    def load_one_disc(self):
        
        self.curr_col = openpyxl.utils.get_column_letter(self.curr_col_num) # Excel column
        formula_df = self.get_formulae()
        
        Excel_Misc_Fns.df_to_worksheet(formula_df[[self.contract]], self.ws_lead, 
                                       index=False, header=True, 
                                       startrow = 7, startcol = self.curr_col_num)
        
    def load_all_disc(self):
        
        num_contracts = len(self.contracts)
        
        if num_contracts > 1:
            self.ws_lead.insert_cols(self.total_col, amount=num_contracts-1)
            self.total_col = self.total_col + num_contracts - 1
        
        for contract_key, contract_value in self.contracts.items():
            
            self.contract = contract_key
            self.contract_value = contract_value
            self.load_one_disc()
            
            self.curr_col_num += 1
            if self.curr_col_num == self.total_col:
                self.curr_col_num -= 1
                break
            
        # Total Col
        curr_col = self.curr_col
        total_col = openpyxl.utils.get_column_letter(self.total_col)
        i = 7
        for r in self.ws_lead[total_col][i:110]:
            if AllDisclosure.VARIABLE_NAME[i-7] != 'blank':
                r.value = Excel_Misc_Fns.sum_formula("B" + str(i+1), curr_col + str(i+1))
            i += 1
            
    def format_all_disc(self):
        
        self.ws_lead['A1'].value = self.client
        
        self.ws_lead['A2'].value = self.ws_lead['A2'].value.replace('<FY END>', 
                                                                    self.fy_end.strftime('%d %B %Y'))
        
        # Fill <PREV FY> and <FY> values.
        for r in self.ws_lead['A'][8:120]:
            if r.value is None:
                pass
            else:
                if "<PREV FY START>" in r.value:
                    r.value = r.value.replace("<PREV FY START>", self.pfy_start.strftime("%d %B %Y"))
                elif "<PREV FY END>" in r.value:
                    r.value = r.value.replace("<PREV FY END>", self.pfy_end.strftime("%d %B %Y"))
                elif "<FY END>" in r.value:
                    r.value = r.value.replace("<FY END>", self.fy_end.strftime("%d %B %Y"))
                elif "<<PREV FY>>" in r.value:
                    r.value = r.value.replace("<<PREV FY>>", str(self.pfy_end.year))
                elif "<<FY>>" in r.value:
                    r.value = r.value.replace("<<FY>>", str(self.fy_end.year))
                    
        # Format header
        for c in self.ws_lead[7]:
            c.font = Font(name='Arial', size=10, bold=True)
            c.alignment = Alignment(horizontal='center', vertical='center', wrapText=True)
    

#%% Tester
if __name__ == '__main__':
    if 1:
    
        import inputs
        
        input_fp = r"D:\Ben\_Ref\Audit DA Curriculum\Module\frs116_automation\autolease_hm_cw\INPUT Q&M 2021\INPUT TEMPLATE - FINAL edited.xlsx"
        output_fp  = r"D:\Ben\_Ref\Audit DA Curriculum\Module\frs116_automation\autolease_hm_cw\INPUT Q&M 2021\OUTPUT.xlsx"
        
        input_fp = r"D:\Ben\_Ref\Audit DA Curriculum\Module\frs116_automation\autolease_hm_cw\AOXIN 2021\INPUT TEMPLATE - Aoxin (26.01.2022) edited.xlsx"
        output_fp = r"D:\Ben\_Ref\Audit DA Curriculum\Module\frs116_automation\autolease_hm_cw\AOXIN 2021\test.xlsx"

        lease_data_reader = inputs.LeaseDataReader(input_fp, sheet_name = 'Lease Data')
        lease_data_reader.__main__()
        
        df = lease_data_reader.df.copy()
        contract = '35'
        pfy_start = datetime.date(2020, 1, 1)
        fy_start = datetime.date(2021,1,1)
        fy_end = pd.to_datetime('2021-12-31 00:00:00')
        
        print(contract)
        
        # self = OneContract(df, contract, fy_start, fy_end, pfy_start)
        # self = OneContractSchedule(df, contract, fy_start, fy_end, pfy_start, output_fp)
        self = OneContractDisclosure(df, contract, fy_start, fy_end, pfy_start, output_fp)
        self.__main__()
        
        # self = OneContractdf, contract, fy_start, fy_end, pfy_start, output_fp,
