# -*- coding: utf-8 -*-
"""
Created on Sun Aug 13 07:28:01 2017

@author: dwlan
"""
import numpy as np
import pandas as pd
import landrys_piping_bom_reader_helper as helper

CATEGORY_VALVE = 'VALVE'
CATEGORY_SUPPORT = 'SUPPORT'
CATEGORY_GASKET = 'GASKET'
CATEGORY_BOLT = 'BOLT'
CATEGORY_FITTING = 'FITTING'
CATEGORY_INSTRUMENT = 'INSTRUMENT'
CATEGORY_SPECIALTY_ITEM = 'SPECIALTY'
CATEGORY_MISC = 'MISC'
CATEGORY_PIPE = 'PIPE'
CATEGORY_FLANGE = 'FLANGE'
CATEGORY_UNCLASSIFIED = 'UNCLASSIFIED'

COLUMN_NAME_QTY = 'qty'
COLUMN_NAME_SIZE = 'size'
COLUMN_NAME_DESCRIPTION = 'description'
COLUMN_NAME_CATEGORY = 'category'
COLUMN_NAME_SUBCATEGORY = 'sub_category'
COLUMN_NAME_QTY_AS_FLOAT = 'qty_as_float'
COLUMN_NAME_SIZE_LIST = 'size_list'
COLUMN_NAME_SIZE_DELEGATE = 'size_delegate'

class bom():
    
    
    def __init__(self, xlsx_filepath):
        self.source_filepath = xlsx_filepath
        self.df = pd.read_excel(xlsx_filepath)
        first_3_col = self.df.columns[0:3].tolist()
        self.df.rename(columns={
                first_3_col[0]: COLUMN_NAME_QTY
                , first_3_col[1]: COLUMN_NAME_SIZE
                , first_3_col[2]: COLUMN_NAME_DESCRIPTION}, inplace=True)
        self.df.columns = [x.lower() for x in self.df.columns]    
        self.remove_blank_records(COLUMN_NAME_DESCRIPTION)
        self.remove_blank_records(COLUMN_NAME_QTY)
        self.add_column(COLUMN_NAME_CATEGORY, '')
        self.add_column(COLUMN_NAME_SUBCATEGORY, '')
        self.add_qty_as_floats_column()
        self.add_size_lists_column()
#%%
        self.classify_end_connections()
        self.classify_ratings()
        self.classify_metallurgy()
        self.classify_schedule()
#%%
        self.butt_welds = []
        self.socket_welds = []
        self.olets = []
        self.threaded_connections = []
        self.bolt_ups = []
#%%
        self.classify_valves()
        self.classify_supports()
        self.classify_gaskets()
        self.classify_bolts()
        self.classify_fittings()
        #self.classify_welds()
        self.classify_instruments()
        self.classify_specialty_items()
        self.classify_misc_items()
        self.classify_pipe()
        self.classify_flanges()
        self.classify_unknowns()
#%%
        self.df_connections = pd.DataFrame(data=None, columns=self.df.columns)
        self.df_connections['connection_type'] = ''
        self.df_connections['connection_qty'] = np.NaN
        self.df_connections['schedule'] = ''
        self.classify_end_connections_for_valves()
        self.classify_end_connections_for_fittings()
        self.classify_end_connections_for_flanges()
#%%
        self.results_filepath=''
    
#%%
        '''
        self.summarize_pipe_handling()
        self.summarize_valve_handling()
        self.summarize_boltups()
        self.summarize_buttwelds()
        self.summarize_socketwelds()
        self.summarize_olet_welds()
        self.summarize_threaded_connections()
        self.summarize_instrument_handling()
        self.summarize_supports()
        self.summarize_spring_cans()
        '''        
        
#%%        
    def add_column(self, column_name, default_value=np.NaN):
        if column_name not in self.df.columns:
            self.df[column_name] = default_value
#%%            
    def remove_blank_records(self, column_name):
        print(str.format("Before blank {0} removed: {1}", 
                         column_name, str(self.df.size)))
        desc_is_null = pd.isnull(self.df[column_name])
        self.df = self.df[:][desc_is_null == False]
        print(str.format("After blank {0} removed: {1}", 
                         column_name, str(self.df.size)))
 #%%       
    def add_qty_as_floats_column(self):
        
        self.df[COLUMN_NAME_QTY_AS_FLOAT]=pd.to_numeric(self.df[COLUMN_NAME_QTY], errors='coerce')
        
        try:
            has_no_values = pd.isnull(self.df[COLUMN_NAME_QTY_AS_FLOAT])
            contains_feet_or_inches = self.df[has_no_values][COLUMN_NAME_QTY].str.contains('["\']')
            contains_feet_or_inches_nulls = pd.isnull(contains_feet_or_inches)
            df_feet_inches_no_nulls = contains_feet_or_inches[:]\
                [contains_feet_or_inches_nulls == False]
            qty_ft_inches = self.df.loc[df_feet_inches_no_nulls.index]
            qty_ft_inches[COLUMN_NAME_QTY_AS_FLOAT] = qty_ft_inches[COLUMN_NAME_QTY]\
                .apply(helper.convert_feet_and_inches_text_to_numerical_feet)
            self.df.loc[qty_ft_inches.index,COLUMN_NAME_QTY_AS_FLOAT] = \
                qty_ft_inches[COLUMN_NAME_QTY_AS_FLOAT]
        except:
            pass
        
#%%    
    def add_size_lists_column(self):
        self.df[COLUMN_NAME_SIZE_LIST] = np.NaN
        self.df[COLUMN_NAME_SIZE_DELEGATE] = np.NaN
        self.df[COLUMN_NAME_SIZE_LIST] = self.df[COLUMN_NAME_SIZE]\
            .apply(helper.parse_multiple_sizes)
        self.df[COLUMN_NAME_SIZE_DELEGATE] = self.df[COLUMN_NAME_SIZE_LIST].apply(lambda s: max(s))
#%%            
    def classify_valves(self):
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, '(?<!PLE-)GATE[ ,]', CATEGORY_VALVE, 'GATE')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, '^(?=.*VA).*CONTROL', CATEGORY_VALVE, 'CONTROL')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, '^(?=.*VA).*GLOBE', CATEGORY_VALVE, 'GLOBE')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, '^(?=.*VA).*CHECK', CATEGORY_VALVE, 'CHECK')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, '^(?=.*VA).*RELEIF|RELIEF', CATEGORY_VALVE, 'RELEIF')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, '^(?=.*VA).*BALL', CATEGORY_VALVE, 'BALL')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, '^(?=.*VA).*BUTT', CATEGORY_VALVE, 'BUTTERFLY')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, '^(?=.*VA).*PLUG', CATEGORY_VALVE, 'PLUG')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, '^(?=.*VA).*NEEDLE', CATEGORY_VALVE, 'NEEDLE')
#%%
    def classify_supports(self):
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'TRUNNION', CATEGORY_SUPPORT, 'DUMMY LEG (TRUNNION)')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'SHOE', CATEGORY_SUPPORT, 'SHOE')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'ROD|HANGER', CATEGORY_SUPPORT, 'ROD HANGER')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'GUIDE', CATEGORY_SUPPORT, 'GUIDE')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'STANCHION', CATEGORY_SUPPORT, 'STANCHION')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'BRACKET', CATEGORY_SUPPORT, 'BRACKET')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'SLEEVE', CATEGORY_SUPPORT, 'SLEEVE')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'SPRING', CATEGORY_SUPPORT, 'SPRING CAN')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'D.MMY', CATEGORY_SUPPORT, 'DUMMY LEG')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'STOP', CATEGORY_SUPPORT, 'STOP')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'BASE', CATEGORY_SUPPORT, 'BASE')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'ANGLE', CATEGORY_SUPPORT, 'ANGLE')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, '^(?!.*ORIFICE).*PLATE', CATEGORY_SUPPORT, 'BASE (PLATE)')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'SUPPORT', CATEGORY_SUPPORT, 'PIPE')
#%%
    def classify_gaskets(self):
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, '^(?!.*FLANG|FLNG)(.*GSK|.*GASKET)', CATEGORY_GASKET)
#%%
    def classify_bolts(self):
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'BOLT', CATEGORY_BOLT)
#%%
    def classify_fittings(self):
        # Fittings with zero connections
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'NIPPLE|NIP,', CATEGORY_FITTING, 'NIPPLE')
            # A stub-end is a part that goes with a lap-joint flange.  
            # Slip-on flanges have two welds associated with it
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'STUB END', CATEGORY_FITTING, 'STUB END')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'SPECTACLE', CATEGORY_FITTING, 'SPECTACLE BLIND')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'PADDLE', CATEGORY_FITTING, 'PADDLE BLIND')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'FIGURE 8', CATEGORY_FITTING, 'FIGURE 8 BLIND')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'BLIND', CATEGORY_FITTING, 'BLIND')
        
        # Fittings with one connection
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'CAP(?!S)', CATEGORY_FITTING, 'CAP')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'PLUG', CATEGORY_FITTING, 'PLUG')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'OLET|O-LET|[WELTSF]-O-L|[WTSF]OL ', CATEGORY_FITTING, 'OLET')
            # The smaller connection end of the olet is captured by the fitting or valve
            # that it is connected to.
        
        # Fittings with two connections
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, '45 [DE]|45D |90 |90D |ELBOW|\WELL|^ELL', CATEGORY_FITTING, 'ELL')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'COUPLING|^(?!.*PIPE).*CPLG|.*CLPG|.*COUPL(?!ED)', CATEGORY_FITTING, 'COUPLING')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'REDUC|SWAGE|SWG|CONC[\., ]|ECC SW|ECC R|RED ECC|ECC', CATEGORY_FITTING, 'REDUCER')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'UNION', CATEGORY_FITTING, 'UNION')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'OFFSET', CATEGORY_FITTING, 'OFFSET')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'PIPET', CATEGORY_FITTING, 'PIPET')
        
        
        # Fittings with three connections
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, '(?<!\w)TEE', CATEGORY_FITTING, 'TEE')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'LATERAL', CATEGORY_FITTING, 'LATERAL TEE')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'WYE', CATEGORY_FITTING, 'WYE')
        
        
        # Fittings with four connections
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'CROSS', CATEGORY_FITTING, 'CROSS')
#%%
    def classify_instruments(self):
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'PRESSURE INDIC|(?<!A)PI-',CATEGORY_INSTRUMENT, 'PRESSURE INDICATOR')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'PRESSURE DIF|PT-|PIT-|PDIT-|PT-', CATEGORY_INSTRUMENT, 'PRESSURE TRANSMITTER')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, '[LTFPX]V-', CATEGORY_INSTRUMENT, 'CONTROL VALVE')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'FE-|^(?!.*FLG|.*FLAN).*ORIFICE', CATEGORY_INSTRUMENT, 'FLOW ELEMENT')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'LT-', CATEGORY_INSTRUMENT, 'LEVEL TRANSMITTER')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'TW-', CATEGORY_INSTRUMENT, 'THERMOWELL')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'TI-', CATEGORY_INSTRUMENT, 'TEMPERATURE INDICATOR')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'TE-', CATEGORY_INSTRUMENT, 'TEMPERATURE ELEMENT')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'AE-|ANALYZER', CATEGORY_INSTRUMENT, 'ANALYZER ELEMENT')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'INSTRUMENT', CATEGORY_INSTRUMENT)
#%%
    def classify_specialty_items(self):
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'STEAM|STM TRAP', CATEGORY_SPECIALTY_ITEM, 'STEAM TRAP')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'SP-', CATEGORY_SPECIALTY_ITEM)
        
#%%
    def classify_misc_items(self):
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'CHAIN OP', CATEGORY_MISC, 'CHAIN OPERATOR')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'BLEED R', CATEGORY_MISC, 'BLEED RING')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'RUPTURE', CATEGORY_MISC, 'RUPTURE DISK')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'DISTRIBUTOR', CATEGORY_MISC, 'DISTRIBUTOR')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'PROBE', CATEGORY_MISC, 'PROBE')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'REINFORCING PAD', CATEGORY_MISC, 'REINFORCING PAD')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'HOSE', CATEGORY_MISC, 'HOSE')
        #helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'TAP', CATEGORY_MISC, 'TAP CONNECTION')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'STRAINER', CATEGORY_MISC, 'STRAINER')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'MIXER', CATEGORY_MISC, 'MIXER')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'TRANSLATIONAL', CATEGORY_MISC, 'TRANSLATIONAL')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'SS STATION', CATEGORY_MISC, 'SS STATION')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'SINGLE BELL', CATEGORY_MISC, 'SINGLE BELL')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, 'PIPE PG', CATEGORY_MISC, 'PIPE GUIDE')
#%%
    def classify_pipe(self):
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, CATEGORY_PIPE, 'PIPE')
#%%
    def classify_flanges(self):
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, '^(?=.*FLG|.*FLAN)(.*THR|.*THD)', CATEGORY_FLANGE, 'THREADED')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, '^(?=.*FLG|.*FLAN).*WN', CATEGORY_FLANGE, 'WELD-NECK')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, '^(?=.*FLG|.*FLAN).*LAP[ -]|.*LJ', CATEGORY_FLANGE, 'LAP-JOINT')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, '^(?=.*FLG|.*FLAN)(.*SO|.*SLIP)', CATEGORY_FLANGE, 'SLIP-ON')
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, '^(?=.*FLG|.*FLAN).*SW', CATEGORY_FLANGE, 'SOCKET-WELD')
#%%
    def classify_unknowns(self):
        helper.apply_classification(self.df, COLUMN_NAME_DESCRIPTION, '', CATEGORY_UNCLASSIFIED)
#%% 
    # The connection details associated to SW's are:
    # - size
    # - rating (3000#, 6000#, 9000#)
    # - metallurgy     
    # The connection details associated with BW's are:
    # - size
    # - schedule
    # - metallurgy  
    def classify_end_connections(self):
        helper.apply_end_connection_classification(
                self.df, COLUMN_NAME_DESCRIPTION, 
                'THR|THD|SCRD|SCREW|FSE|MSE|NPT|T[BLOS]E', 'THD')
        helper.apply_end_connection_classification(
                self.df, COLUMN_NAME_DESCRIPTION, 
                'SW(?![AG])|SO(?!LID)|SLIP|P.E|xPE|\bPE\b|PEx', 'SW')
        helper.apply_end_connection_classification(
                self.df, COLUMN_NAME_DESCRIPTION, 
                'BW[ .,]|LAP|LJ|WN|(?<![PT])BE|B.E|^(?!.*P.E).*WPB', 'BW')
#%%
    def classify_end_connections_for_valves(self):
        #Should only need to classify connections for SW valves or Threaded valves.  Flanged
        #Flanged valves have no connection associated with them - only valve handling and bolt-up
        df_valves = self.df[self.df[COLUMN_NAME_CATEGORY] == CATEGORY_VALVE]
        df_3way_valves = df_valves[df_valves['description'].str.contains('[3E].{0,3}WAY', case=False)]
        df_sw_valves = df_valves[df_valves['connection_type_list'].apply(lambda s: 'SW' in s)]
        #df_bw_valves = df_valves[df_valves['connection_type_list'].apply(lambda s: 'BW' in s)]
        df_thd_valves = df_valves[df_valves['connection_type_list'].apply(lambda s: 'THD' in s)]
        df_reducing_valves = df_valves[df_valves[COLUMN_NAME_SIZE_LIST].apply(lambda s: len(s)>1)]
        
        df_valves_with_one_connection_type = df_valves[df_valves['connection_type_list'].apply(lambda s: len(s) == 1)]
        
        
        
        pass
#%%
    def classify_end_connections_for_fittings(self):
        #Don't forget to add Seal Welds for B
        df = self.df[self.df[COLUMN_NAME_CATEGORY] == CATEGORY_FITTING]
        pass
#%%
    def classify_end_connections_for_flanges(self):
        df = self.df[self.df[COLUMN_NAME_CATEGORY] == CATEGORY_FLANGE]
        pass
#%%
    def classify_ratings(self):    
        helper.apply_rating_classification(self.df, COLUMN_NAME_DESCRIPTION, '2500', 'CL 2500')
        helper.apply_rating_classification(self.df, COLUMN_NAME_DESCRIPTION, '1500', 'CL 1500')
        helper.apply_rating_classification(self.df, COLUMN_NAME_DESCRIPTION, '900(?!0)', 'CL 900')
        helper.apply_rating_classification(self.df, COLUMN_NAME_DESCRIPTION, '800', 'CL 800')
        helper.apply_rating_classification(self.df, COLUMN_NAME_DESCRIPTION, '600(?!0)', 'CL 600')
        helper.apply_rating_classification(self.df, COLUMN_NAME_DESCRIPTION, '400', 'CL 400')
        helper.apply_rating_classification(self.df, COLUMN_NAME_DESCRIPTION, '300(?!0)', 'CL 300')
        helper.apply_rating_classification(self.df, COLUMN_NAME_DESCRIPTION, '250', 'CL 250')
        helper.apply_rating_classification(self.df, COLUMN_NAME_DESCRIPTION, '200', 'CL 200')
        helper.apply_rating_classification(self.df, COLUMN_NAME_DESCRIPTION, '175', 'CL 175')
        helper.apply_rating_classification(self.df, COLUMN_NAME_DESCRIPTION, '150', 'CL 150')
        helper.apply_rating_classification(self.df, COLUMN_NAME_DESCRIPTION, '125', 'CL 125')
        helper.apply_rating_classification(self.df, COLUMN_NAME_DESCRIPTION, '3000', '3000#')
        helper.apply_rating_classification(self.df, COLUMN_NAME_DESCRIPTION, '6000', '6000#')
        helper.apply_rating_classification(self.df, COLUMN_NAME_DESCRIPTION, '9000', '9000#')
#%%
    def classify_metallurgy(self):
        # CS metallurgy
        helper.apply_metallurgy_classification(self.df, COLUMN_NAME_DESCRIPTION, 'CS|A10[65]|A-10[56]|WP[ABC]|WC[ABC]', 'CS', 'CARBON STEEL')
        
        # Alloy metallurgy
        # ref https://www.engineeringtoolbox.com/astm-fittings-materials-d_248.html
        #    for cross reference of metallurgy specifications
        #helper.apply_metallurgy_classification(self.df, COLUMN_NAME_DESCRIPTION, 'C[rR]|ALLOY|A197|A338|316|A536|A403|INCOLOY', 'ALLOY')
        helper.apply_metallurgy_classification(self.df, COLUMN_NAME_DESCRIPTION, '316L', 'ALLOY', '316L SS')
        helper.apply_metallurgy_classification(self.df, COLUMN_NAME_DESCRIPTION, '317L', 'ALLOY', '317L SS')
        helper.apply_metallurgy_classification(self.df, COLUMN_NAME_DESCRIPTION, '304L', 'ALLOY', '304L SS')
        helper.apply_metallurgy_classification(self.df, COLUMN_NAME_DESCRIPTION, '(?=.*316)(?=.*SS).*', 'ALLOY', '316 SS')
        helper.apply_metallurgy_classification(self.df, COLUMN_NAME_DESCRIPTION, '(?=.*317)(?=.*SS).*', 'ALLOY', '317 SS')
        helper.apply_metallurgy_classification(self.df, COLUMN_NAME_DESCRIPTION, '(?=.*304)(?=.*SS).*', 'ALLOY', '304 SS')
        helper.apply_metallurgy_classification(self.df, COLUMN_NAME_DESCRIPTION, '(?=.*321)(?=.*SS).*', 'ALLOY', '321 SS')
        
        helper.apply_metallurgy_classification(self.df, COLUMN_NAME_DESCRIPTION, '(A.{0,1}182.{0,7}F.{0,1}51)', 'ALLOY', 'A182 F51 DUPLEX SS')
        
        helper.apply_metallurgy_classification(self.df, COLUMN_NAME_DESCRIPTION, '1.{0,10}C[Rr].{0,2}1/2.{0,2}M[Oo]', 'ALLOY', '1 1/4Cr-1/2Mo')
        helper.apply_metallurgy_classification(self.df, COLUMN_NAME_DESCRIPTION, '(A.{0,1}182.{0,7}F.{0,1}11)', 'ALLOY', '1 1/4Cr-1/2Mo')
        helper.apply_metallurgy_classification(self.df, COLUMN_NAME_DESCRIPTION, '(A.{0,1}217.{0,8}WC6)', 'ALLOY', '1 1/4Cr-1/2Mo')
        
        helper.apply_metallurgy_classification(self.df, COLUMN_NAME_DESCRIPTION, '9.{0,1}C[Rr].{0,2}1 {0,1}M[Oo]', 'ALLOY', '9Cr-1Mo')
        helper.apply_metallurgy_classification(self.df, COLUMN_NAME_DESCRIPTION, '(A.{0,1}182.{0,7}F.{0,1}9)', 'ALLOY', '9Cr-1Mo')
        helper.apply_metallurgy_classification(self.df, COLUMN_NAME_DESCRIPTION, '(A.{0,1}217.{0,8}C12)', 'ALLOY', '9Cr-1Mo')
        
        helper.apply_metallurgy_classification(self.df, COLUMN_NAME_DESCRIPTION, '5.{0,1}C[Rr].{0,2}1/2 {0,1}M[Oo]', 'ALLOY', '5Cr-1/2Mo')
        helper.apply_metallurgy_classification(self.df, COLUMN_NAME_DESCRIPTION, '(A.{0,1}182.{0,7}F.{0,1}5)', 'ALLOY', '5Cr-1/2Mo')
        
        helper.apply_metallurgy_classification(self.df, COLUMN_NAME_DESCRIPTION, '(?=.*2205)(?=.*DUPLEX).*', 'ALLOY', '2205 DUPLEX')
        helper.apply_metallurgy_classification(self.df, COLUMN_NAME_DESCRIPTION, 'C276', 'ALLOY', 'HASTELLOY C276')
        
        helper.apply_metallurgy_classification(self.df, COLUMN_NAME_DESCRIPTION, '(PVC)(HDPE)', 'PVC', 'PVC')
        
        helper.apply_metallurgy_classification(self.df, COLUMN_NAME_DESCRIPTION, 'BRONZ', 'BRONZE', 'BRONZE')
        
        helper.apply_metallurgy_classification(self.df, COLUMN_NAME_DESCRIPTION, 'A.{0,1}197|GALV', 'GALVANIZED', 'GALVANIZED')
#%%
    def classify_schedule(self):
        helper.apply_schedule_classification(self.df, COLUMN_NAME_DESCRIPTION, '(SC{0,1}H|S/)(?=.{0,7}40S)', 'SCH 40S') 
        helper.apply_schedule_classification(self.df, COLUMN_NAME_DESCRIPTION, '(?=.*STD)', 'SCH STD')
        helper.apply_schedule_classification(self.df, COLUMN_NAME_DESCRIPTION, '(SC{0,1}H|S/)(?=.{0,7}[Hx\s\.]40(,|x|\s|$))', 'SCH 40')
        helper.apply_schedule_classification(self.df, COLUMN_NAME_DESCRIPTION, '(SC{0,1}H|S/)(?=.{0,7}10S)', 'SCH 10S')
        helper.apply_schedule_classification(self.df, COLUMN_NAME_DESCRIPTION, '(SC{0,1}H|S/)(?=.{0,7}[Hx\s\.]10(,|x|\s|$))', 'SCH 10')
        helper.apply_schedule_classification(self.df, COLUMN_NAME_DESCRIPTION, '(SC{0,1}H|S/)(?=(.{0,7}[Hx\s\.]60|60)(,|x|\s|$))', 'SCH 60')
        helper.apply_schedule_classification(self.df, COLUMN_NAME_DESCRIPTION, '(SC{0,1}H|S/)(?=.{0,7}80S)', 'SCH 80S')
        helper.apply_schedule_classification(self.df, COLUMN_NAME_DESCRIPTION, '(SC{0,1}H|S/)(?=.{0,7}80(,|x|\s|$))', 'SCH 80')
        helper.apply_schedule_classification(self.df, COLUMN_NAME_DESCRIPTION, '(?=.*(?<!X)XS)', 'SCH XS')
        helper.apply_schedule_classification(self.df, COLUMN_NAME_DESCRIPTION, '(SC{0,1}H|S/)(?=.{0,7}100(,|x|\s|$))', 'SCH 100')
        helper.apply_schedule_classification(self.df, COLUMN_NAME_DESCRIPTION, '(SC{0,1}H|S/)(?=.{0,7}120(,|x|\s|$))', 'SCH 120')
        helper.apply_schedule_classification(self.df, COLUMN_NAME_DESCRIPTION, '(SC{0,1}H|S/)(?=.{0,7}140(,|x|\s|$))', 'SCH 140')
        helper.apply_schedule_classification(self.df, COLUMN_NAME_DESCRIPTION, '(SC{0,1}H|S/)(?=.{0,7}160(,|x|\s|$))', 'SCH 160')
        helper.apply_schedule_classification(self.df, COLUMN_NAME_DESCRIPTION, '(?=.*XXS)', 'SCH XXS')
        helper.apply_schedule_classification(self.df, COLUMN_NAME_DESCRIPTION, '(SC{0,1}H|S/)(?=.{0,7}5S)', 'SCH 5S')
        helper.apply_schedule_classification(self.df, COLUMN_NAME_DESCRIPTION, '(SC{0,1}H|S/)(?=.{0,7}[Hx\s\.]5(,|x|\s|$))', 'SCH 5')
        helper.apply_schedule_classification(self.df, COLUMN_NAME_DESCRIPTION, '(SC{0,1}H|S/)(?=.{0,7}[Hx\s\.]20(,|x|\s|$))', 'SCH 20')
        helper.apply_schedule_classification(self.df, COLUMN_NAME_DESCRIPTION, '(SC{0,1}H|S/)(?=.{0,7}[Hx\s\.]30(,|x|\s|$))', 'SCH 30')
        
        
#%%
    def pivot_table_for_pipe_handling(self):
        df_pipe = self.df[self.df[COLUMN_NAME_CATEGORY] == CATEGORY_PIPE]
        
        schedule = df_pipe['schedule_list'].apply(lambda s: helper.get_first_value_from_list(s))
        df_pipe = df_pipe.assign(schedule=schedule)
        pvt = pd.pivot_table(data = df_pipe, 
                             index = ['size_delegate', 'schedule'], 
                             values = 'qty_as_float',
                             aggfunc = np.sum,
                             fill_value = 0)
        return pvt
#%%
    def pph(self):
        return self.pivot_table_for_pipe_handling()
        
        
#%%
    def pivot_table_for_valve_handling(self):
        df_valves = self.df[self.df[COLUMN_NAME_CATEGORY] == CATEGORY_VALVE]
        
        # ADD LOGIC:
        #    ONLY INCLUDE VALVES THAT ARE BOLT-UP CONNECTED
        pvt = pd.pivot_table(data = df_valves, 
                             index = ['size_delegate', 'rating'], 
                             values = 'qty_as_float',
                             aggfunc = np.sum,
                             fill_value = 0)
        return pvt
#%%
    def pvh(self):
        return self.pivot_table_for_valve_handling()
#%%
    def pivot_table_for_bolt_ups(self):
        df_gaskets = self.df[self.df[COLUMN_NAME_CATEGORY] == CATEGORY_GASKET]
        pvt = pd.pivot_table(data = df_gaskets,
                             index = ['size_delegate', 'rating'],
                             values = 'qty_as_float',
                             aggfunc = np.sum,
                             fill_value = 0)
        return pvt
#%%
    def pivot_table_for_supports(self):
        df_supports = self.df[self.df[COLUMN_NAME_CATEGORY] == CATEGORY_SUPPORT]
        pvt = pd.pivot_table(data = df_supports,
                             index = 'sub_category',
                             values = 'qty_as_float',
                             aggfunc = np.sum,
                             fill_value = 0)
        return pvt
#%%
    def pivot_table_for_buttwelds(self):
        #metallurgy, size, schedule
        
        pass
#%%
    def pivot_table_for_socketwelds(self):
        # metallurgy, size, rating
        
        pass
#%%
    def pivot_table_for_olet_welds(self):
        # metallurgy, size, rating
        
        pass
#%%
    def pivot_table_for_threaded_connections(self):
        # count of quantity
        
        pass
#%%
# http://pbpython.com/pandas-pivot-report.html
    def to_excel(self, filepath = '', open_after_save = True):
        from win32com.client import Dispatch
        
        if not filepath and not self.results_filepath:
            self.results_filepath = self.source_filepath.replace('.xlsx','_results.xlsx')
            
        writer = pd.ExcelWriter(self.results_filepath)
        self.df.to_excel(writer,'bom', freeze_panes=(1,1))
        #for category in self.df['category'].unique():
        #    self.df[self.df["category"]==category].to_excel(writer, category, freeze_panes=(1,1))
        self.pivot_table_for_pipe_handling().to_excel(writer, 'pipe_handling')
        self.pivot_table_for_valve_handling().to_excel(writer, 'valve_handling')
        self.pivot_table_for_bolt_ups().to_excel(writer, 'boltups')
        self.pivot_table_for_supports().to_excel(writer, 'supports')
        '''
        df_pipe = self.df[self.df[COLUMN_NAME_CATEGORY] == CATEGORY_PIPE]
        pipe_table = pd.pivot_table(df_pipe,index=["metallurgy_general","size_delegate"], 
                                    values=["qty_as_float"],aggfunc=[np.sum],fill_value=0)
        for metallurgy in pipe_table.index.get_level_values(0).unique():
            temp_df = pipe_table.xs(metallurgy, level=0)
            temp_df.to_excel(writer, metallurgy)
        '''
        writer.save()
        
        if open_after_save:
            xl = Dispatch("Excel.Application")
            xl.Visible = True # otherwise excel is hidden
            
        # newest excel does not accept forward slash in path
            wb = xl.Workbooks.Open(self.results_filepath)
        #wb.Close()
        #xl.Quit()
        
        
        # In[138]:
        
        ''' Use OpenPyXl to further manipulate the excel file'''
        #wkb = xl.load_workbook(results_filepath)
        #print(wkb.get_sheet_names())
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        