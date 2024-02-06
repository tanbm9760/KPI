import pickle
import json
from datetime import datetime
from pathlib import Path
import os.path
import pandas as pd
import streamlit as st
from streamlit_option_menu import option_menu
from streamlit_extras.stylable_container import stylable_container
from streamlit_extras.metric_cards import style_metric_cards
import streamlit_authenticator as stauth
import plotly.express as px
import plotly.graph_objects as go
from PIL import Image
st.set_page_config(page_title="KPI Dashboard WebApp",
                   page_icon=":bar_chart:",
                   layout="wide"
)
# ----USER AUTHENTICATION----   
names = ["Admin"]
usernames = ["Admin"]
# Load hashed passwords
file_path = Path(__file__).parent / "hashed_pw.pkl"
with file_path.open("rb") as file:
    hashed_passwords = pickle.load(file)

authenticator = stauth.Authenticate(names, usernames, hashed_passwords, "kpi_dashboard", "abcdef", cookie_expiry_days=30)

name, authenticator_status, username = authenticator.login("Login","main")
if authenticator_status == False:
    st.error("Username/password is incorrect")

if authenticator_status == None:
    st.warning("Please enter your username and password")

if authenticator_status:
    with st.sidebar:
        authenticator.logout("Logout","sidebar")
        st.title(f"Welcome {name}")
        selected_sites = option_menu(None, ["NNC", "Department",  "Person"],
            icons=['house', 'journal-bookmark', 'person'],
            menu_icon="cast",
            default_index=0,
            orientation="vertical",
            styles={
                "container": {"padding": "0!important", "background-color": "#0b3a75", "min-width": "100%"},
                "icon": {"color": "orange", "font-size": "15px"},
                "nav-link": {"font-size": "15px", "text-align": "left", "margin":"0px", "--hover-color": "#165f9c"},
                "nav-link-selected": {"background-color": "#165f9c"},
            }
        )
        current_month = datetime.now().month
        current_year = datetime.now().year
        if(current_month == 1):
            default_month = 11
            default_year = current_year - 1
        else:
            default_month = current_month - 2
            default_year = current_year
        month_selected = st.sidebar.selectbox(
            "Month?", 
            ['1','2','3','4','5','6','7','8','9','10','11','12'],
            default_month,
            ) 
        year_selected = st.sidebar.selectbox(
            "Year? ",
            list(range(default_year-5,default_year+5)),
            5
        ) 
    # ----MAINPAGE----       
    excel_name = './data/' + str(year_selected) + '-' + str(month_selected) +'.xlsx'
    checkfile = os.path.isfile(excel_name)
    
    #Hàm tạo chỉ số thống kê
    def statistical_index(dataframe):
        #-total-
        total_works_assigned = int(dataframe["Assigned"].sum())
        total_works_urgent = int(dataframe["Urgent"].sum())
        total_works_important = int(dataframe["Important"].sum())
        total_works_completed = int(dataframe["Completed"].sum())
        total_works_late = int(dataframe["Late"].sum())
        total_works_outofdate = int(dataframe["Out of date"].sum())
        #-display-
        metric_1, metric_2, metric_3, metric_4, metric_5, metric_6 = st.columns(6)
        metric_1.metric("Assigned Works",f"{total_works_assigned:,}")
        metric_2.metric("Urgent Works",f"{total_works_urgent:,}")
        metric_3.metric("Important Works",f"{total_works_important:,}")
        metric_4.metric("Completed Works",f"{total_works_completed:,}")
        metric_5.metric("Late Works",f"{total_works_late:,}")
        metric_6.metric("Out of date Works",f"{total_works_outofdate:,}") 
        style_metric_cards(
                background_color='rgba(38,151,215,0.4)',
                border_radius_px=5,
                border_size_px=1,
                border_color="rgb(255,255,255)",
                border_left_color='rgb(255,255,255)',
                box_shadow=True,
            )
    #----BARCHART----
    def works_bar(dataframe,groupby,sortby,value,title,color):
        df_filter = dataframe.groupby(by=[groupby]).sum().sort_values(by=sortby)
        #-title-
        st.write(title)
        #-config-
        fig_work_bar = px.bar(
            df_filter,
            y=value,
            x=df_filter.index,
            height=320,
            text_auto='.0f',
            )
        fig_work_bar.update_traces(
            textposition = 'auto', 
            textfont_size=10,
            marker_color=color,
            )
        fig_work_bar.update_layout(
            {'plot_bgcolor': 'rgba(0, 0, 0, 0)','paper_bgcolor': 'rgba(0, 0, 0, 0)',}
            )
        fig_work_bar.update_layout(
            margin=dict(t=20, r=40, b=0, l=40), 
            showlegend = False, 
            xaxis_title=None,
            yaxis_title=None,
            xaxis_tickangle=90,
            )
        fig_work_bar.update_yaxes(
            rangemode="tozero"
            )
        #-display-
        with stylable_container(
            key="chart",
            css_styles="""
            {   
                border-radius: 0.5em;
                border: 2px groove white;
                box-shadow: rgba(255, 255, 255, 0.5) 0px 2px 4px 0px, rgba(38,151,215,0.4) 0px 2px 12px 0px;
                background-color: rgba(38,151,215,0.4);
            }"""
        ):
            st.plotly_chart(fig_work_bar, use_container_width= True)   
    #----RANKING TOP WORK BY DEPARTMENT----
    def ranktop5high_index(key,css,dfname,colvalue,unit):
        with stylable_container(
                    key=key,
                    css_styles=css
                    ):
                    st.write(str(dfname.nlargest(n=10, columns=colvalue).index[0]),";", str(dfname.nlargest(n=10, columns=colvalue).iloc[0].at[colvalue]), unit)
        with stylable_container(
            key=key,
            css_styles=""
            ):
            st.write(str(dfname.nlargest(n=10, columns=colvalue).index[1]),";", str(dfname.nlargest(n=10, columns=colvalue).iloc[1].at[colvalue]), unit)
        with stylable_container(
            key=key,
            css_styles=""
            ):
            st.write(str(dfname.nlargest(n=10, columns=colvalue).index[2]),";", str(dfname.nlargest(n=10, columns=colvalue).iloc[2].at[colvalue]), unit)
        with stylable_container(
            key=key,
            css_styles=""
            ):
            st.write(str(dfname.nlargest(n=10, columns=colvalue).index[3]),";", str(dfname.nlargest(n=10, columns=colvalue).iloc[3].at[colvalue]), unit)
        with stylable_container(
            key=key,
            css_styles=""
            ):
            st.write(str(dfname.nlargest(n=10, columns=colvalue).index[4]),";", str(dfname.nlargest(n=10, columns=colvalue).iloc[4].at[colvalue]), unit)
    def ranktop5high_person(key,css,dfname,colvalue,colif1,colif2,unit):
        with stylable_container(
                    key=key,
                    css_styles=css
                    ):
                    st.write(str(dfname.nlargest(n=10, columns=colvalue).iloc[0].at[colif1]),";",str(dfname.nlargest(n=10, columns=colvalue).iloc[0].at[colif2]),";", str(round(dfname.nlargest(n=10, columns=colvalue).iloc[0].at[colvalue], 1)), unit)
        with stylable_container(
            key=key,
            css_styles=""
            ):
            st.write(str(dfname.nlargest(n=10, columns=colvalue).iloc[1].at[colif1]),";",str(dfname.nlargest(n=10, columns=colvalue).iloc[1].at[colif2]),";", str(round(dfname.nlargest(n=10, columns=colvalue).iloc[1].at[colvalue], 1)), unit)
        with stylable_container(
            key=key,
            css_styles=""
            ):
            st.write(str(dfname.nlargest(n=10, columns=colvalue).iloc[2].at[colif1]),";",str(dfname.nlargest(n=10, columns=colvalue).iloc[2].at[colif2]),";", str(round(dfname.nlargest(n=10, columns=colvalue).iloc[2].at[colvalue], 1)), unit)
        with stylable_container(
            key=key,
            css_styles=""
            ):
            st.write(str(dfname.nlargest(n=10, columns=colvalue).iloc[3].at[colif1]),";",str(dfname.nlargest(n=10, columns=colvalue).iloc[3].at[colif2]),";", str(round(dfname.nlargest(n=10, columns=colvalue).iloc[3].at[colvalue], 1)), unit)
        with stylable_container(
            key=key,
            css_styles=""
            ):
            st.write(str(dfname.nlargest(n=10, columns=colvalue).iloc[4].at[colif1]),";",str(dfname.nlargest(n=10, columns=colvalue).iloc[4].at[colif2]),";", str(round(dfname.nlargest(n=10, columns=colvalue).iloc[4].at[colvalue], 1)), unit)
    def ranktop5low_person(key,css,dfname,colvalue,colif1,colif2,unit):
        with stylable_container(
                    key=key,
                    css_styles=css
                    ):
                    st.write(str(dfname.nsmallest(n=10, columns=colvalue).iloc[0].at[colif1]),";",str(dfname.nsmallest(n=10, columns=colvalue).iloc[0].at[colif2]),";", str(round(dfname.nsmallest(n=10, columns=colvalue).iloc[0].at[colvalue], 1)), unit)
        with stylable_container(
            key=key,
            css_styles=""
            ):
            st.write(str(dfname.nsmallest(n=10, columns=colvalue).iloc[1].at[colif1]),";",str(dfname.nsmallest(n=10, columns=colvalue).iloc[1].at[colif2]),";", str(round(dfname.nsmallest(n=10, columns=colvalue).iloc[1].at[colvalue], 1)), unit)
        with stylable_container(
            key=key,
            css_styles=""
            ):
            st.write(str(dfname.nsmallest(n=10, columns=colvalue).iloc[2].at[colif1]),";",str(dfname.nsmallest(n=10, columns=colvalue).iloc[2].at[colif2]),";", str(round(dfname.nsmallest(n=10, columns=colvalue).iloc[2].at[colvalue], 1)), unit)
        with stylable_container(
            key=key,
            css_styles=""
            ):
            st.write(str(dfname.nsmallest(n=10, columns=colvalue).iloc[3].at[colif1]),";",str(dfname.nsmallest(n=10, columns=colvalue).iloc[3].at[colif2]),";", str(round(dfname.nsmallest(n=10, columns=colvalue).iloc[3].at[colvalue], 1)), unit)
        with stylable_container(
            key=key,
            css_styles=""
            ):
            st.write(str(dfname.nsmallest(n=10, columns=colvalue).iloc[4].at[colif1]),";",str(dfname.nsmallest(n=10, columns=colvalue).iloc[4].at[colif2]),";", str(round(dfname.nsmallest(n=10, columns=colvalue).iloc[4].at[colvalue], 1)), unit)
    def Works_LeaderBoards(dataframe):
        ranking_1, ranking_2, ranking_3, ranking_4, ranking_5 = st.columns(5)
        with ranking_1:
            with stylable_container(
                key="rank",
                css_styles="""
                {   
                    padding: 0.5em 0.1em;
                    border-radius: 0.5em;
                    text-align: center;
                    background-color: rgba(38,151,215,0.4);
                    border: 2px groove white;
                    box-shadow: rgba(255, 255, 255, 0.5) 0px 2px 4px 0px, rgba(38,151,215,0.4) 0px 2px 12px 0px;
                }"""
            ):
                st.subheader("Urgent")
                ranktop5high_person(
                    "r_t1",
                    """{border-radius: 0.5em;                    
                        background-color: rgba(38,151,215,0.7);
                        text-align: center;}""",
                    dataframe,
                    'Urgent',
                    'ID HRM',
                    'Name',
                    '',
                )
        with ranking_2:
            with stylable_container(
                key="rank",
                css_styles=""
            ):
                st.subheader("Important")
                ranktop5high_person(
                    "r_t2",
                    """{border-radius: 0.5em;                    
                        background-color: rgba(255,127,14,0.7);
                        text-align: center;}""",
                    dataframe,
                    'Important',
                    'ID HRM',
                    'Name',
                    '',
                )
        with ranking_3:
            with stylable_container(
                key="rank",
                css_styles=""
            ):
                st.subheader("Completed")
                ranktop5high_person(
                    "r_t3",
                    """{border-radius: 0.5em;                    
                        background-color: rgba(44,160,44,0.7);
                        text-align: center;}""",
                    dataframe,
                    'Completed',
                    'ID HRM',
                    'Name',
                    '',
                )
        with ranking_4:
            with stylable_container(
                key="rank",
                css_styles=""
            ):
                st.subheader("Late")
                ranktop5high_person(
                    "r_t4",
                    """{border-radius: 0.5em;                    
                        background-color: rgba(214,39,40,0.7);
                        text-align: center;}""",
                    dataframe,
                    'Late',
                    'ID HRM',
                    'Name',
                    '',
                )
        with ranking_5:
            with stylable_container(
                key="rank",
                css_styles=""
            ):
                st.subheader("Out of Date")
                ranktop5high_person(
                    "r_t5",
                    """{border-radius: 0.5em;                    
                        background-color: rgba(148,103,189,0.7);
                        text-align: center;}""",
                    dataframe,
                    'Out of date',
                    'ID HRM',
                    'Name',
                    '',
                )
    #------------------------------------------------------------------
    def fig_coef_his(dataframe):
        #-config-
        fig_coef_his = px.histogram(
            dataframe,
            x="Coefficient", 
            color='Department',
            height= 350, 
            text_auto='.0f',
            color_discrete_sequence=px.colors.qualitative.D3,  
            )
        fig_coef_his.update_layout(                   
            {'plot_bgcolor': 'rgba(0, 0, 0, 0)','paper_bgcolor': 'rgba(0, 0, 0, 0)'}
        )
        fig_coef_his.update_layout(bargap=0.1,xaxis_title='KPI Coefficient', yaxis_title='Total')
        fig_coef_his.update_xaxes(dtick=0.1)
        fig_coef_his.update_traces(
            hoverinfo='x+y', textposition = 'inside', textfont_size=10, textangle = 0,
        ) 
        #-title-
        st.write("KPI Coefficient")
        #-display-
        with stylable_container(
            key="chart",
            css_styles=""
        ):
            st.plotly_chart(fig_coef_his, use_container_width= True)
    def fig_kpifinal_his(dataframe):
        #-config-
        fig_kpifinal_his = px.histogram(
            dataframe, 
            x="Kpi Final",
            color='Department',
            color_discrete_sequence=px.colors.qualitative.Light24, 
            height= 350, 
            text_auto='.0f',
            )
        fig_kpifinal_his.update_layout(                   
            {'plot_bgcolor': 'rgba(0, 0, 0, 0)','paper_bgcolor': 'rgba(0, 0, 0, 0)'}
        )
        fig_kpifinal_his.update_layout(bargap=0.1,xaxis_title='KPI Final Score', yaxis_title='Total')
        fig_kpifinal_his.update_xaxes(dtick=5)
        fig_kpifinal_his.update_traces(
            hoverinfo='x+y', textposition = 'inside', textfont_size=10, textangle = 0,
        )
        #-title-
        st.write("KPI Final Score")
        #-display-
        with stylable_container(
            key="chart",
            css_styles=""
        ):
            st.plotly_chart(fig_kpifinal_his, use_container_width= True)
    def KPI_Leaderboards(dataframe):
        rankkpi_1, rankkpi_2, rankkpi_3, rankkpi_4, rankkpi_5 = st.columns(5)
        with rankkpi_1:
            with stylable_container(
                key="rankkpi",
                css_styles="""
                {   
                    border-radius: 0.5em;
                    padding: 0.5em 0.1em;  
                    text-align: center;
                    border: 2px groove white;
                    box-shadow: rgba(255, 255, 255, 0.5) 0px 2px 4px 0px, rgba(38,151,215,0.4) 0px 2px 12px 0px;
                    background-color: rgba(38,151,215,0.4);
                }"""
            ):
                st.subheader("KPI Works")
                ranktop5high_person(
                    "rk_t1",
                    """{border-radius: 0.5em;
                        padding: 0.2em 0em;                   
                        background-color: rgba(38,151,215,0.7);
                        text-align: center;}""",
                    dataframe,
                    'Kpi Works',
                    'ID HRM',
                    'Name',
                    '',
                )
        with rankkpi_2:
            with stylable_container(
                key="rankkpi",
                css_styles=""
            ):
                st.subheader("KPI Highest Bonus")
                ranktop5high_person(
                    "rk_t2",
                    """{border-radius: 0.5em;
                        padding: 0.2em 0em;                       
                        background-color: rgba(255,127,14,0.7);
                        text-align: center;}""",
                    dataframe,
                    'Kpi Bonus',
                    'ID HRM',
                    'Name',
                    '',
                )
        with rankkpi_3:
            with stylable_container(
                key="rankkpi",
                css_styles=""
            ):
                st.subheader("KPI Lowest Bonus")
                ranktop5low_person(
                    "rk_t3",
                    """{border-radius: 0.5em;   
                        padding: 0.2em 0em;                    
                        background-color: rgba(214,39,40,0.7);
                        text-align: center;}""",
                    dataframe,
                    'Kpi Bonus',
                    'ID HRM',
                    'Name',
                    '',
                )
        with rankkpi_4:
            with stylable_container(
                key="rankkpi",
                css_styles=""
            ):
                st.subheader("KPI Final")
                ranktop5high_person(
                    "rk_t4",
                    """{border-radius: 0.5em; 
                        padding: 0.2em 0em;                      
                        background-color: rgba(44,160,44,0.7);
                        text-align: center;}""",
                    dataframe,
                    'Kpi Final',
                    'ID HRM',
                    'Name',
                    '',
                )
        with rankkpi_5:
            with stylable_container(
                key="rankkpi",
                css_styles=""
            ):
                st.subheader("KPI Coefficient")
                ranktop5high_person(
                    "rk_t5",
                    """{border-radius: 0.5em;  
                        padding: 0.2em 0em;                     
                        background-color: rgba(148,103,189,0.7);
                        text-align: center;}""",
                    dataframe,
                    'Coefficient',
                    'ID HRM',
                    'Name',
                    '',
                )
    def highlight_columns(col):
        color = 'rgba(255,127,14,0.7)' if col.name in ['Kpi Final', 'Coefficient'] else 'rgba(11,58,117,0.7)' 
        return ['background-color: {}'.format(color) for _ in col]
    ###################################################################
    if checkfile == True:
        df = pd.read_excel(io=excel_name, 
                    engine='openpyxl', 
                    sheet_name='Data',  
                    skiprows=3,
                    usecols='B:Q',
                    nrows=1000,
        )
        if selected_sites == "NNC":
            df_selection = df
        # -------------------------------------------------------------
            with stylable_container(
                    key="title",
                    css_styles="""
                    {   
                        text-align: center;
                    }"""
                ):
                st.header("3C .Inc Key Performance Indicators Dashboard") 
                st.header("🗃️ Data " + str(month_selected) + "/" + str(year_selected)+' 🗃️')  
            st.markdown("---")
            # ---------------------------------------------------------
            st.subheader("1️⃣ Statistical Index")
            st.markdown("###")
            statistical_index(df_selection)
            st.markdown("---")
            # ---------------------------------------------------------
            st.subheader("2️⃣ Work LeaderBoards")
            st.markdown("###")
            Works_LeaderBoards(df_selection)
            st.markdown("---")
            # ---------------------------------------------------------
            st.subheader("3️⃣ Works Histogram")
            st.markdown("###")
            fig_11, fig_12, fig_13 = st.columns(3)
            with fig_11:
                works_bar(df_selection, "Department", "Department", "Assigned", "Assigned Works", "rgb(220,220,220)")
                works_bar(df_selection, "Department", "Department", "Completed", "Completed Works", "rgb(44,160,44)")
            with fig_12:
                works_bar(df_selection, "Department", "Department", "Urgent", "Urgent Works", "rgb(38,151,215)")
                works_bar(df_selection, "Department", "Department", "Late", "Late Works", "rgb(214,39,40)")
            with fig_13:
                works_bar(df_selection, "Department", "Department", "Important", "Important Works", "rgb(255,127,14)")
                works_bar(df_selection, "Department", "Department", "Out of date", "Out of date Works", "rgb(148,103,189)")
            st.markdown("---")
            # ------------------------------------------------------------------------------
            st.subheader("4️⃣ KPI LeaderBoards")
            st.markdown("###")
            KPI_Leaderboards(df_selection)
            st.markdown("---")
            # ------------------------------------------------------------------------------
            st.subheader("5️⃣ KPI Histogram")
            st.markdown("###")
            fig_21, fig_22 = st.columns([2,1])
            with fig_21:
                fig_kpifinal_his(df_selection)
            with fig_22:
                fig_coef_his(df_selection)
            st.markdown("---")
            # ------------------------------------------------------------------------------
            st.subheader("6️⃣ Data Frame")
            st.markdown("###")
            styled_df = df_selection.style.apply(highlight_columns)
            st.dataframe(styled_df,width=1400)
        if selected_sites == "Department" :
            department = st.sidebar.multiselect(
                "Select Department",
                options=df["Department"].unique(),
                max_selections=3,
            )
            df_selection = df.query(
                "Department == @department"
            )
            # ------------------------------------------------------------------------------
            with stylable_container(
                    key="title",
                    css_styles="""
                    {   
                        text-align: center;
                    }"""
                ):
                st.header("3C .Inc Key Performance Indicators Dashboard") 
                st.header("🗃️ Data " + str(month_selected) + "/" + str(year_selected)+' 🗃️') 
            st.markdown("---")
            # ------------------------------------------------------------------------------
            if department:
                st.subheader("1️⃣ Statistical Index")
                st.markdown("###")
                statistical_index(df_selection)
                st.markdown("---")
                # ------------------------------------------------------------------------------
                st.subheader("2️⃣ Work LeaderBoards")
                st.markdown("###")
                Works_LeaderBoards(df_selection)
                st.markdown("---")
                # ------------------------------------------------------------------------------
                st.subheader("3️⃣ Works Histogram")
                st.markdown("###")
                fig_11, fig_12, fig_13 = st.columns(3)
                with fig_11:
                    works_bar(df_selection, "Name", "Department", "Assigned", "Assigned Works", "rgb(220,220,220)")
                    works_bar(df_selection, "Name", "Department", "Completed", "Completed Works", "rgb(44,160,44)")
                with fig_12:
                    works_bar(df_selection, "Name", "Department", "Urgent", "Urgent Works", "rgb(38,151,215)")
                    works_bar(df_selection, "Name", "Department", "Late", "Late Works", "rgb(214,39,40)")
                with fig_13:
                    works_bar(df_selection, "Name", "Department", "Important", "Important Works", "rgb(255,127,14)")
                    works_bar(df_selection, "Name", "Department", "Out of date", "Out of date Works", "rgb(148,103,189)")
                st.markdown("---")
            else:
                st.write("Please Select Department!")
        if selected_sites == "Person" :
            department = st.sidebar.multiselect(
                "Select Department",
                options=df["Department"].unique(),
            )
            df_buffer = df.query(
                "Department == @department",
                )
            name = st.sidebar.multiselect(
                "Select Person",
                options = df_buffer["Name"].unique(),
            )
            df_selection = df_buffer.query(
                "Name == @name",
                )
        # -----------------------------------------------------------------------------------
            with stylable_container(
                    key="title",
                    css_styles="""
                    {   
                        text-align: center;
                    }"""
                ):
                st.header("3C .Inc Key Performance Indicators Dashboard") 
                st.header("🗃️ Data " + str(month_selected) + "/" + str(year_selected)+' 🗃️')  

            st.markdown("---")
        # -----------------------------------------------------------------------------------
            st.subheader("1️⃣ Statistical Index")
            st.markdown("###")
            statistical_index(df_selection)
            st.markdown("---")
        # -----------------------------------------------------------------------------------
    else:
        st.write("No data")