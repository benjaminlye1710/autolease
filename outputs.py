### Import from dependencies
from dependencies import *
import inputs, engines

# =============================================================================
#### Output class
# =============================================================================
class LeaseLiabilityWriter:
    '''Writes final disclosure and individual contract schedules to specified filepath (Excel).'''
    
    def __init__(self, input_fp, output_fp):
        
        self.input_fp = input_fp
    
        self.output_fp = output_fp
        
        self.__main__()
        
    def __main__(self):
        
        self.create_output_wb(overwrite = False)
        self.read_all_contracts()
        self.write_all_contracts()
        self.write_all_disclosure()
        
        
    def create_output_wb(self, overwrite = False):
        
        wb = openpyxl.load_workbook(self.input_fp)
        
        if not overwrite and os.path.exists(self.output_fp):
            
            folder_name = os.path.dirname(self.output_fp)
            
            [file_name, extension] = \
                os.path.basename(self.output_fp).split('.')
            
            now = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            
            output_fp_timestamped = \
                os.path.join(folder_name, (f'{file_name}_{now}.{extension}'))
            
            self.output_fp = output_fp_timestamped
        
        wb.save(self.output_fp)
        
        wb.close()
    
    
    def read_all_contracts(self):
        
        self.lease_data = inputs.LeaseDataReader(self.input_fp, sheet_name = 'Lease Data',)
        self.lease_data.__main__()
    
    
    def write_all_contracts(self):
        
        lease_data = self.lease_data
        contracts = {}
        
        df_gb = lease_data.df.groupby('Contract')
        
        for contract in df_gb.groups.keys():
            
            one_contract_disclosure = (
                engines.OneContractDisclosure(
                    lease_data.df, contract, lease_data.fy_start, 
                    lease_data.fy_end, lease_data.pfy_start, self.output_fp)
                )
            
            one_contract_disclosure.__main__()
            
            contracts[contract] = {
                'input_df': one_contract_disclosure.df,
                'schedule_df': one_contract_disclosure.schedule_df,
                'disclosure_df': one_contract_disclosure.disc_df
                }
        
        self.contracts = contracts.copy()
        
        
    def write_all_disclosure(self):
        
        lease_data = self.lease_data
        contracts = self.contracts.copy()
        self.all_disclosure = (
            engines.AllDisclosure(
                lease_data.client, contracts, self.output_fp, 
                lease_data.pfy_start, lease_data.pfy_end, lease_data.fy_end)
            )
        self.all_disclosure.__main__()

            

#%% Tester
if __name__ == "__main__":
    
    
    #%%% BEN'S TESTER
    if 0:
        
        #%%%% 
        if 0:
            # mocked cases
            input_fp = r"D:\Ben\_Ref\Audit DA Curriculum\Module\frs116_automation\autolease_hm_cw\INPUT TEMPLATE 2021\INPUT TEMPLATE.xlsx"
            output_fp = r"D:\Ben\_Ref\Audit DA Curriculum\Module\frs116_automation\autolease_hm_cw\INPUT TEMPLATE 2021\OUTPUT.xlsx"
        if 1:
            input_fp = r"D:\Ben\_Ref\Audit DA Curriculum\Module\frs116_automation\autolease_hm_cw\INPUT Q&M 2020\INPUT TEMPLATE.xlsx"
            output_fp  = r"D:\Ben\_Ref\Audit DA Curriculum\Module\frs116_automation\autolease_hm_cw\INPUT Q&M 2020\OUTPUT.xlsx"
        
        self = LeaseLiabilityWriter(input_fp, output_fp)
        lease_liability_writer = self

    #%%%%         
    if 0: 
        #%%%%  DEBUG
        pass
    

    
    #%%% ANIKA'S TESTER
    if 1:
        
        #%%%% 

        # input_fp = r"D:\Documents\autolease-main\FRS 116 - Workings List co - sample output_anonymised INPUT.xlsx"
        # output_fp = r'D:\Documents\autolease-main\FRS 116 - Workings List co - sample output_anonymised OUTPUT v' + f'{datetime.date.today().strftime("%Y%m%d")}' + '.xlsx'
        
        input_fp = "D:\Documents\Lease Liability\FRS 116 - Workings List co - sample output_anonymised.xlsx"
        output_fp = "D:\Documents\Lease Liability\FRS 116 - Workings List co - OUTPUT test.xlsx" 
        
        # input_fp = "D:\Documents\Lease Liability\margaret\FRS 116 - Workings List co - sample output_anonymised_Version_SentToEngagementTeam.xlsx"
        # output_fp = "D:/Documents/Lease Liability/margaret/FRS 116 - Workings List co - sample output_anonymised_Version_SentToEngagementTeam OUTPUT.xlsx"
        
        run = LeaseLiabilityWriter(input_fp,output_fp)

