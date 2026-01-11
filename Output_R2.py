import streamlit as st
import pandas as pd
import numpy as nm
import time

#############styling table#################
def style(df):
    return (
        df.style
        # ---------- base cell style ----------
        .set_properties(**{
            "font-size": "20px",
            "text-align": "center",
            "color": "#C7F2FB",
            "background-color": "#111111D3",
            "border": "1px solid #d1d5db"
        })
        )

#############styling table#################

st.set_page_config(page_title="Reporting Tool", layout="wide")
with open("style.css") as f:
    st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

st.markdown(
    """
    <style>
        [data-testid="stSidebar"] {
            max-width: 410px;
            min-width: 410px;
        }
    </style>
    """,
    unsafe_allow_html=True
)
with st.sidebar:
    st.image("download2.png", width=260)
    st.markdown("### 📊 Reporting Tool   © GRK")
    st.markdown("#### Input Section")
    min_minutes = st.number_input("Minimum minutes for SAIDI-SAIFI", value=5, min_value=0)
    max_minutes = st.number_input("Maximum minutes for SAIDI-SAIFI", value=120, min_value=10)

    consumerdata = st.file_uploader("Upload main consumer data for SAIDI-SAIFI", type=["xls","xlsx"])
    uploaded_file = st.file_uploader("Upload your HIS data", type=["xlsx", "xls"])

    run_clicked = st.button("RUN")
    st.write("---")
    st.markdown("### Contact")
    st.write("Galibur Rahman Khan Niloy")
    st.write("Assistant Engineer")
    st.write("GSP&EA(SCADA)")


st.markdown("<h1 style='text-align: center; font-size:60px; font-weight: bold;'>SCADA REPORTING TOOL</h1>", unsafe_allow_html=True)
st.header("*REPORTS*")


if run_clicked:
    if uploaded_file is not None:
        with st.spinner("Processing your data... Please wait ⏳"):
            time.sleep(5)       # simulate processing
        st.success("Report generated successfully! Please wait")
        # Choose engine based on extension
        if uploaded_file.name.endswith(".xlsx"):
            data = pd.read_excel(uploaded_file)
        else:
            data = pd.read_excel(uploaded_file, engine="xlrd")
        df= data[data['Element'].str.contains(r'Circuit Breaker|Master Trip|OC Relay Operated|EF Relay Operated|NonDir OC50Stg1 Optd|Non-Dir EF Operated|Non-Dir EF 50N Stg2|OC or EF Rly Trip|VCB Trip on Fault|OvrCurr Prot. Stg2|Eth Flt Prot. Stg2|Eth Flt Prot. Stg1|General Prot. Trip|Transfrmer Diff Prot|Non-Dir EF 50N Stg1|OvrCurr Prot. Stg1|Oil Temperature Trip|Sens. EF Operated|OC & EF Operated|General Prot Trip|REF Protection|PRD Trip (Main Tank)|27 Stg1 UV Trip|Winding Temp Trip|Standby EF Trip|Protection Tripped|59 Stage1 OV trip|Trxf Main Prt Trip|Non-Dir OC Operated|51 OC|Transfrmer Diff Prot|Oil Temperature Trip|Trxf Main Prt Trip|Winding Temp Trip|Protection Tripped|REF Protection|PRD Trip (Main Tank)', case=False, na=False)]
        df = df[df['A'].isna()]
        df = df.sort_values(by=['B1', 'B2', 'B3', 'Time stamp', 'Milliseconds'],ascending=[True, True, True, True, True]).reset_index(drop=True)
        TripSignal =["Master Trip", "OC Relay Operated", "EF Relay Operated", "NonDir OC50Stg1 Optd", "Non-Dir EF Operated", "Non-Dir EF 50N Stg2", "OC or EF Rly Trip", "VCB Trip on Fault", "OvrCurr Prot. Stg2", "Eth Flt Prot. Stg2", "Eth Flt Prot. Stg1", "General Prot. Trip", "Transfrmer Diff Prot", "Non-Dir EF 50N Stg1", "OvrCurr Prot. Stg1", "Oil Temperature Trip", "Sens. EF Operated", "OC & EF Operated", "General Prot Trip", "REF Protection", "PRD Trip (Main Tank)", "27 Stg1 UV Trip", "Winding Temp Trip", "Standby EF Trip", "Protection Tripped", "59 Stage1 OV trip", "Trxf Main Prt Trip", "Non-Dir OC Operated", "51 OC", "Transfrmer Diff Prot", "Oil Temperature Trip", "Trxf Main Prt Trip", "Winding Temp Trip", "Protection Tripped","REF Protection","PRD Trip (Main Tank)"]
        ###################################tripcount###################################################
        condition = (
            ((df["Element"]== "Circuit Breaker") & (df["Status"] == "open")) & 
            (
                (df["B3"] == df["B3"].shift(1)) |
                (df["B3"] == df["B3"].shift(-1))
            ) &
            (
                ((df["Element"].shift(1).isin(TripSignal)) & ((df["Time stamp"].shift(1) - df["Time stamp"]).dt.total_seconds().abs() <= 3)) |   # Previous row is Trip
                ((df["Element"].shift(-1).isin(TripSignal)) & ((df["Time stamp"] - df["Time stamp"].shift(-1)).dt.total_seconds().abs() <= 3))  # Next row is Trip
            )
        )
        result = df[condition]
        df_tr = result[result['B3'].str.contains('_TR', case=True, na=False) &
                       ~result['B3'].str.contains(r'_TR_S', na=False) ]

        # Non-TR rows
        df_non_tr = result[~(result['B3'].str.contains('_TR', case=True, na=False) &
                        ~result['B3'].str.contains(r'_TR_S', na=False) )]

        st.write(f"Number of Feeder Trip: {len(df_non_tr)}")  # Row count
        st.write(f"Number of Transformer Trip: {len(df_tr)}") 

        count_11_t = (df_tr["B2"].astype(str) == "11").sum()
        count_33_t = (df_tr["B2"].astype(str) == "33").sum()
        count_132_t = (df_tr["B2"].astype(str) == "132").sum()

        
        count_11_f = (df_non_tr["B2"].astype(str) == "11").sum()
        count_33_f = (df_non_tr["B2"].astype(str) == "33").sum()
        count_132_f = (df_non_tr["B2"].astype(str) == "132").sum()

        col1, col2, col3 = st.columns(3)

        col1.header("11KV")
        col2.header("33KV")
        col3.header("132KV")

        col1.write(f" Feeder Trip: {count_11_f}")
        col1.write(f"Transformer Trip: {count_11_t}")
        col2.write(f" Feeder Trip: {count_33_f}")
        col2.write(f"Transformer Trip: {count_33_t}")
        col3.write(f" Feeder Trip: {count_132_f}")
        col3.write(f"Transformer Trip: {count_132_t}")
        ######Tabs
        tab1,tab2,tab3,tab4,tab5,tab6,tab7 =st.tabs(['Feeder Trip Count','Transformer Trip Count','Trip Summary','Trip Summary with Trials','Consumer No.','SAIDI_SAIFI(Trip Only)','SAIDI-SAIFI(ALL)'])
        ######Tabs
        output = df_non_tr[['B1','B2','B3']]
        output1=output.sort_values(by=['B1','B2','B3'])
        tripcount = output1.groupby(['B1', 'B2', 'B3']).size().reset_index(name='No. of Trip')
        
        with tab1:
            st.dataframe(style(tripcount),use_container_width=True,height=600)
            #st.write(df)
        output = df_tr[['B1','B2','B3']]
        output1=output.sort_values(by=['B1','B2','B3'])
        tripcount = output1.groupby(['B1', 'B2', 'B3']).size().reset_index(name='No. of Trip')
        with tab2:
            st.dataframe(style(tripcount),use_container_width=True,height=600)
    #################################################tripcount##########################################################

    #################################################tripsummary########################################################
        condition2 = (
            ((df["Element"]== "Circuit Breaker") & (df["Status"] == "open")) & 
            (
                (df["B3"] == df["B3"].shift(1)) |
                (df["B3"] == df["B3"].shift(-1))
            ) &
            (
                ((df["Element"].shift(1).isin(TripSignal)) & ((df["Time stamp"].shift(1) - df["Time stamp"]).dt.total_seconds().abs() <= 30)) |   # Previous row is Trip
                ((df["Element"].shift(-1).isin(TripSignal)) & ((df["Time stamp"] - df["Time stamp"].shift(-1)).dt.total_seconds().abs() <= 3))  # Next row is Trip
            ) |
            ((df["Element"]== "Circuit Breaker") & (df["Status"] == "close"))
        )
        tripsummary = df[condition2]
        tripsummary['Time stamp'] = pd.to_datetime(tripsummary['Time stamp'])
        tripsummary1= tripsummary[['Time stamp', 'B1','B2', 'B3', 'Status']]
        tripsummary222=tripsummary1.sort_values(by=['B1','B2', 'B3', 'Time stamp']).reset_index(drop=True) #original trip summary with trials

        ##############condition to exclude trials############
        condition_close = (
            (tripsummary222['Status'] == 'close') &
            (tripsummary222['Status'].shift(-1) == 'open') & (tripsummary222['B3'] == tripsummary222['B3'].shift(-1)) &
            ((tripsummary222['Time stamp'].shift(-1) - tripsummary222['Time stamp']).dt.total_seconds() < 2)
            )
        # Mark the next "open" row
        condition_open = condition_close.shift(1, fill_value=False)

        # Combine both
        condition_both = condition_close | condition_open
            # delete those rows

        tripsummary2 = tripsummary222[~condition_both].reset_index(drop=True)
        trial = tripsummary222[condition_both].reset_index(drop=True)

        #####################################################with trials####################
        tripsummary222['Close Timestamp'] = pd.NaT

        for idx, row in tripsummary222.iterrows():
        # everything inside the loop must be indented
            if (row['Status'] == 'open'
            and tripsummary222.loc[idx+1, 'Status'] == 'close' 
            and row['B3'] == tripsummary222.loc[idx+1, 'B3'] ):
                tripsummary222.loc[idx, 'Close Timestamp'] = tripsummary222.loc[idx+1, 'Time stamp']
        
        tripsummary333 = tripsummary222[tripsummary222['Status'] == 'open'].reset_index(drop=True)

        tripsummary333 = tripsummary333.rename(columns={
        'B3': 'Feeder Name',
        'B1': 'SubStation',
        'B2': 'Voltage',
        'Time stamp': 'Trip Time',
        'Close Timestamp': 'Close Time'
        })
        tripsummary333['Duration'] = pd.NaT
        for idx, row in tripsummary333.iterrows():
                tripsummary333.loc[idx, 'Duration'] = tripsummary333.loc[idx, 'Close Time'] - tripsummary333.loc[idx, 'Trip Time']
                tripsummary333['Duration'] = tripsummary333['Duration'].apply(lambda x: str(x).split('.')[0])
        TRS333=tripsummary333[['SubStation','Voltage','Feeder Name', 'Trip Time', 'Close Time','Duration']] 

##################################################################################
        tripsummary2['Close Timestamp'] = pd.NaT

        for idx, row in tripsummary2.iterrows():
        # everything inside the loop must be indented
            if (row['Status'] == 'open'
            and tripsummary2.loc[idx+1, 'Status'] == 'close' 
            and row['B3'] == tripsummary2.loc[idx+1, 'B3'] ):
                tripsummary2.loc[idx, 'Close Timestamp'] = tripsummary2.loc[idx+1, 'Time stamp']
        
        tripsummary3 = tripsummary2[tripsummary2['Status'] == 'open'].reset_index(drop=True)

        month_end_2359 = (
        tripsummary3['Time stamp']
        .dt.to_period('M')
        .dt.to_timestamp('M')
        + pd.Timedelta(hours=23, minutes=59)
        )

        tripsummary3.loc[tripsummary3['Close Timestamp'].isna(), 'Close Timestamp'] = month_end_2359

        tripsummary3 = tripsummary3.rename(columns={
        'B3': 'Feeder Name',
        'B1': 'SubStation',
        'B2': 'Voltage',
        'Time stamp': 'Trip Time',
        'Close Timestamp': 'Close Time'
        })
        tripsummary3['Duration'] = pd.NaT
        for idx, row in tripsummary3.iterrows():
                tripsummary3.loc[idx, 'Duration'] = tripsummary3.loc[idx, 'Close Time'] - tripsummary3.loc[idx, 'Trip Time']
                tripsummary3['Duration'] = tripsummary3['Duration'].apply(lambda x: str(x).split('.')[0])
        TRS=tripsummary3[['SubStation','Voltage','Feeder Name', 'Trip Time', 'Close Time','Duration']]
        #tripsummary3['Duration'] = tripsummary3['Duration'].apply(lambda x: str(x).split('.')[0])
        with tab3:
            st.dataframe(style(TRS),use_container_width=True,height=600) 
        with tab4:
            st.dataframe(style(TRS333),use_container_width=True,height=600)
################################################tripsummary#########################################################

################################################SAIDI_SAIFI#########################################################

        TRS = TRS[['SubStation', 'Voltage', 'Feeder Name', 'Duration']]
        TRS['Duration'] = pd.to_timedelta(TRS['Duration'])
        TRS['Duration'] = pd.to_numeric(TRS['Duration'].dt.total_seconds() / 60, errors='coerce')
        TRS['Duration'] = TRS['Duration'].fillna(0).astype(int)

        TRS_non_tr = TRS[~TRS['Feeder Name'].str.contains('_TR', case=True, na=False)]
        
        if consumerdata is not None:
        # Choose engine based on extension
            if uploaded_file.name.endswith(".xlsx"):
                consumer_df = pd.read_excel(consumerdata, engine="openpyxl")
            else:
                consumer_df = pd.read_excel(consumerdata, engine="xlrd")

        TRS_non_tr= TRS_non_tr[
            (TRS_non_tr['Duration'] >= min_minutes) 
            ]
        TRS_non_tr['Duration'] = TRS_non_tr['Duration'].clip(upper= max_minutes)

        #TRS_non_tr = TRS_non_tr.groupby(['SubStation','Voltage','Feeder Name'])['Duration'].sum().reset_index()
        TRS_non_tr = (TRS_non_tr.groupby(['SubStation','Voltage','Feeder Name']).agg(Duration=('Duration', 'sum'),TripNo=('Duration', 'count')).reset_index())   
        with tab5:
            st.dataframe(style(consumer_df),use_container_width=True,height=600)
        
        consumer_df = consumer_df.rename(columns={
        'Feeder': 'Feeder Name',
        'Substation': 'SubStation',
        })

        consumer_df= consumer_df[['SubStation', 'Voltage', 'Feeder Name', 'ConsumerNo', 'S&D']]

        SS= consumer_df.merge(TRS_non_tr, how='left', on= 'Feeder Name')
        SS = SS.rename(columns={
        'Voltage_x': 'Voltage',
        'SubStation_x': 'SubStation',
        })
        SS=SS[['SubStation', 'Voltage', 'Feeder Name', 'S&D','ConsumerNo', 'Duration', 'TripNo']]
        SS['ConsumerMinute'] = pd.NA
        SS['Consumer_freq'] =pd.NA
        SS['SAIDI'] = pd.NA
        SS['SAIFI'] =pd.NA
        SS['ConsumerMinute'] = SS['ConsumerNo'] * SS['Duration']
        SS['Consumer_freq']= SS['ConsumerNo']*SS['TripNo'] 
        saidi_value = SS['ConsumerMinute'].sum() / SS['ConsumerNo'].sum()
        saifi_value = SS['Consumer_freq'].sum()/ SS['ConsumerNo'].sum()
        SS.loc[0, 'SAIDI'] = saidi_value
        SS.loc[0, 'SAIFI'] = saifi_value

        with tab6:
            st.dataframe(style(SS),use_container_width=True,height=600)
################################################################################################################################################################################
################################################################################################################################################################################
        df=df[((df['Priority']==2) | (df['Status']== 'close'))]
        df_all_saidisaifi = df[['B1','B2','B3','Element','Status','Time stamp','Milliseconds']]
        df_all_saidisaifi = df_all_saidisaifi[df_all_saidisaifi['Element'] == 'Circuit Breaker']
        df_all_saidisaifi['Close Timestamp'] = pd.NaT
        df_all_saidisaifi['Time stamp'] = pd.to_datetime(df_all_saidisaifi['Time stamp'])
        df_all_saidisaifi=df_all_saidisaifi.sort_values(by=['B1','B2', 'B3', 'Time stamp','Milliseconds']).reset_index(drop=True)

        for idx, row in df_all_saidisaifi.iterrows():
            if (row['Status'] == 'open' and df_all_saidisaifi.loc[idx+1, 'Status'] == 'close' and row['B3']== df_all_saidisaifi.loc[idx+1, 'B3']):
                df_all_saidisaifi.loc[idx, 'Close Timestamp'] = df_all_saidisaifi.loc[idx+1, 'Time stamp']
        df_all_saidisaifi = df_all_saidisaifi[df_all_saidisaifi['B2'] == '11']
        df_all_saidisaifi = df_all_saidisaifi[~df_all_saidisaifi['B3'].str.contains('_TR', case=True, na=False)]
        df_all_saidisaifi = df_all_saidisaifi[df_all_saidisaifi['Close Timestamp'].notna()]

        df_all_saidisaifi = df_all_saidisaifi.rename(columns={
        'B3': 'Feeder Name',
        'B1': 'SubStation',
        'B2': 'Voltage',
        'Time stamp': 'Open Time',
        'Close Timestamp': 'Close Time'
        })
        df_all_saidisaifi = df_all_saidisaifi[['SubStation','Voltage','Feeder Name','Open Time','Close Time']]
        df_all_saidisaifi['Duration'] = pd.NaT

        df_all_saidisaifi.loc[df_all_saidisaifi['Close Time'].isna(), 'Close Time'] = month_end_2359

        for idx, row in df_all_saidisaifi.iterrows():
            df_all_saidisaifi.loc[idx, 'Duration'] = (df_all_saidisaifi.loc[idx, 'Close Time'] - df_all_saidisaifi.loc[idx, 'Open Time'])

        df_all_saidisaifi['Duration'] = pd.to_timedelta(df_all_saidisaifi['Duration'])
        df_all_saidisaifi['Duration'] = pd.to_numeric(df_all_saidisaifi['Duration'].dt.total_seconds() / 60, errors='coerce')
        df_all_saidisaifi['Duration'] = df_all_saidisaifi['Duration'].fillna(0).astype(int)

        df_all_saidisaifi= df_all_saidisaifi[
            (df_all_saidisaifi['Duration'] >= min_minutes) 
            ]
        df_all_saidisaifi['Duration'] = df_all_saidisaifi['Duration'].clip(upper= max_minutes) 

        df_all_saidisaifi = (df_all_saidisaifi.groupby(['SubStation','Voltage','Feeder Name']).agg(InteruptionTime=('Duration','sum'),InteruptionNo=('Duration','count')).reset_index()) 
        SS_ALL= SS.merge(df_all_saidisaifi, how='left', on= 'Feeder Name')
        SS_ALL = SS_ALL.rename(columns={
        'Voltage_x': 'Voltage',
        'SubStation_x': 'SubStation',
        'ConsumerMinute': 'Trip_Minute',
        'Consumer_freq': 'Trip_freq'
        })

        SS_ALL = SS_ALL[['SubStation', 'Voltage', 'Feeder Name', 'S&D','ConsumerNo', 'Duration', 'TripNo','Trip_Minute','Trip_freq','SAIDI','SAIFI','InteruptionTime','InteruptionNo']]

        SS_ALL['ALL ConsumerMinute'] = pd.NA
        SS_ALL['ALL Consumer_freq'] =pd.NA
        SS_ALL['With Planned SAIDI'] = pd.NA
        SS_ALL['With Planned SAIFI'] =pd.NA

        SS_ALL['ALL ConsumerMinute'] = SS_ALL['ConsumerNo'] * SS_ALL['InteruptionTime']
        SS_ALL['ALL Consumer_freq']= SS_ALL['ConsumerNo']*SS_ALL['InteruptionNo'] 
        Planned_saidi_value = SS_ALL['ALL ConsumerMinute'].sum() / SS_ALL['ConsumerNo'].sum()
        Planned_saifi_value = SS_ALL['ALL Consumer_freq'].sum()/ SS_ALL['ConsumerNo'].sum()
        SS_ALL.loc[0, 'With Planned SAIDI'] = Planned_saidi_value
        SS_ALL.loc[0, 'With Planned SAIFI'] = Planned_saifi_value

    
        with tab7:
            st.dataframe(style(SS_ALL),use_container_width=True,height=600)






