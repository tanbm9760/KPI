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

#-------------------------------------------------------------------
#---------------------------FUNCTION--------------------------------
#-------------------------------------------------------------------
def save_value(key):
    st.session_state[key] = st.session_state["_"+key]
def get_value(key):
    st.session_state["_"+key] = st.session_state[key] 
#-------------------------------------------------------------------
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
    metric_5.metric("Completed Late Works",f"{total_works_late:,}")
    metric_6.metric("Out of date Works",f"{total_works_outofdate:,}") 
    style_metric_cards(
            background_color='rgba(38,151,215,0.4)',
            border_radius_px=5,
            border_size_px=1,
            border_color="rgb(255,255,255)",
            border_left_color='rgb(255,255,255)',
            box_shadow=True,
        )
#-------------------------------------------------------------------
def works_bar_sum(dataframe,groupby,sortby,value,title,color,barwidth):
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
        width = barwidth,
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
#-------------------------------------------------------------------
def works_bar_mean(dataframe,groupby,sortby,value,title,color,barwidth):
    df_filter = dataframe.groupby(by=[groupby])[value].mean()
    #-title-
    st.write(title)
    #-config-
    fig_work_bar = px.bar(
        df_filter,
        y=value,
        x=df_filter.index,
        height=320,
        text_auto='.1f',
        )
    fig_work_bar.update_traces(
        textposition ='auto', 
        textfont_size=10,
        marker_color=color,
        width = barwidth,
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
#-------------------------------------------------------------------
def ranktop3high_index(key,css,dfname,colvalue,unit):
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
    # with stylable_container(
    #     key=key,
    #     css_styles=""
    #     ):
    #     st.write(str(dfname.nlargest(n=10, columns=colvalue).index[3]),";", str(dfname.nlargest(n=10, columns=colvalue).iloc[3].at[colvalue]), unit)
    # with stylable_container(
    #     key=key,
    #     css_styles=""
    #     ):
    #     st.write(str(dfname.nlargest(n=10, columns=colvalue).index[4]),";", str(dfname.nlargest(n=10, columns=colvalue).iloc[4].at[colvalue]), unit)
#-------------------------------------------------------------------
def ranktop3high_person(key,css,dfname,colvalue,colif1,colif2,unit):
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
    # with stylable_container(
    #     key=key,
    #     css_styles=""
    #     ):
    #     st.write(str(dfname.nlargest(n=10, columns=colvalue).iloc[3].at[colif1]),";",str(dfname.nlargest(n=10, columns=colvalue).iloc[3].at[colif2]),";", str(round(dfname.nlargest(n=10, columns=colvalue).iloc[3].at[colvalue], 1)), unit)
    # with stylable_container(
    #     key=key,
    #     css_styles=""
    #     ):
    #     st.write(str(dfname.nlargest(n=10, columns=colvalue).iloc[4].at[colif1]),";",str(dfname.nlargest(n=10, columns=colvalue).iloc[4].at[colif2]),";", str(round(dfname.nlargest(n=10, columns=colvalue).iloc[4].at[colvalue], 1)), unit)
#-------------------------------------------------------------------
def ranktop3low_person(key,css,dfname,colvalue,colif1,colif2,unit):
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
    # with stylable_container(
    #     key=key,
    #     css_styles=""
    #     ):
    #     st.write(str(dfname.nsmallest(n=10, columns=colvalue).iloc[3].at[colif1]),";",str(dfname.nsmallest(n=10, columns=colvalue).iloc[3].at[colif2]),";", str(round(dfname.nsmallest(n=10, columns=colvalue).iloc[3].at[colvalue], 1)), unit)
    # with stylable_container(
    #     key=key,
    #     css_styles=""
    #     ):
    #     st.write(str(dfname.nsmallest(n=10, columns=colvalue).iloc[4].at[colif1]),";",str(dfname.nsmallest(n=10, columns=colvalue).iloc[4].at[colif2]),";", str(round(dfname.nsmallest(n=10, columns=colvalue).iloc[4].at[colvalue], 1)), unit)
#-------------------------------------------------------------------
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
            st.write("Urgent")
            ranktop3high_person(
                "r_t1",
                """{border-radius: 0.5em;                    
                    background-color: rgba(38,151,215,0.8);
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
            st.write("Important")
            ranktop3high_person(
                "r_t2",
                """{border-radius: 0.5em;                    
                    background-color: rgba(255,127,14,0.8);
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
            st.write("Completed")
            ranktop3high_person(
                "r_t3",
                """{border-radius: 0.5em;                    
                    background-color: rgba(44,160,44,0.8);
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
            st.write("Completed Late")
            ranktop3high_person(
                "r_t4",
                """{border-radius: 0.5em;                    
                    background-color: rgba(214,39,40,0.8);
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
            st.write("Out of Date")
            ranktop3high_person(
                "r_t5",
                """{border-radius: 0.5em;                    
                    background-color: rgba(148,103,189,0.8);
                    text-align: center;}""",
                dataframe,
                'Out of date',
                'ID HRM',
                'Name',
                '',
            )
#-------------------------------------------------------------------
def fig_coef_his(dataframe):
    #-config-
    fig_coef_his = px.histogram(
        dataframe,
        x="Coefficient",
        # color='Department',
        height= 400, 
        text_auto='.0f',
        # color_discrete_sequence=px.colors.qualitative.D3,  
        )
    fig_coef_his.update_layout(                   
        {'plot_bgcolor': 'rgba(0, 0, 0, 0)','paper_bgcolor': 'rgba(0, 0, 0, 0)'}
    )
    fig_coef_his.update_layout(
        bargap=0.1,
        xaxis_title='KPI Coefficient',
        xaxis_tickangle=0,
        yaxis_title=None,
        margin=dict(t=20, r=40, b=0, l=40), 
        showlegend = False, 
        )
    fig_coef_his.update_xaxes(dtick=0.1)
    fig_coef_his.update_traces(
        hoverinfo='x+y', textposition = 'outside', textfont_size=10, textangle = 0, xbins=dict(start=0.05,end=2.5,size=0.1)
    ) 
    #-title-
    st.write("KPI Coefficient")
    #-display-
    with stylable_container(
        key="chart",
        css_styles=""
    ):
        st.plotly_chart(fig_coef_his, use_container_width= True)
#-------------------------------------------------------------------
def fig_kpifinal_his(dataframe):
    #-config-
    fig_kpifinal_his = px.histogram(
        dataframe, 
        x="Kpi Final",
        # color='Department',
        # color_discrete_sequence=px.colors.qualitative.Light24, 
        height= 400, 
        text_auto='.0f',
        )
    fig_kpifinal_his.update_layout(                   
        {'plot_bgcolor': 'rgba(0, 0, 0, 0)','paper_bgcolor': 'rgba(0, 0, 0, 0)'}
    )
    fig_kpifinal_his.update_layout(
        bargap=0.1,
        xaxis_title='KPI Final Score',
        xaxis_tickangle=0,
        yaxis_title=None,
        margin=dict(t=20, r=40, b=0, l=40), 
        showlegend = False,)
    fig_kpifinal_his.update_xaxes(dtick=5)
    fig_kpifinal_his.update_traces(
        hoverinfo='x+y', textposition = 'outside', textfont_size=10, textangle = 0,
    )
    #-title- 
    st.write("KPI Final Score")
    #-display-
    with stylable_container(
        key="chart",
        css_styles=""
    ):
        st.plotly_chart(fig_kpifinal_his, use_container_width= True)
#-------------------------------------------------------------------
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
            st.write("KPI Works")
            ranktop3high_person(
                "rk_t1",
                """{border-radius: 0.5em;
                    padding: 0.2em 0em;                   
                    background-color: rgba(38,151,215,0.8);
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
            st.write("KPI Highest Bonus")
            ranktop3high_person(
                "rk_t2",
                """{border-radius: 0.5em;
                    padding: 0.2em 0em;                       
                    background-color: rgba(255,127,14,0.8);
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
            st.write("KPI Lowest Bonus")
            ranktop3low_person(
                "rk_t3",
                """{border-radius: 0.5em;   
                    padding: 0.2em 0em;                    
                    background-color: rgba(214,39,40,0.8);
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
            st.write("KPI Final")
            ranktop3high_person(
                "rk_t4",
                """{border-radius: 0.5em; 
                    padding: 0.2em 0em;                      
                    background-color: rgba(44,160,44,0.8);
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
            st.write("KPI Coefficient")
            ranktop3high_person(
                "rk_t5",
                """{border-radius: 0.5em;  
                    padding: 0.2em 0em;                     
                    background-color: rgba(148,103,189,0.8);
                    text-align: center;}""",
                dataframe,
                'Coefficient',
                'ID HRM',
                'Name',
                '',
            )
#-------------------------------------------------------------------
def highlight_columns(col):
    color = 'rgba(255,127,14,0.7)' if col.name in ['Kpi Final', 'Coefficient'] else 'rgba(11,58,117,0.7)' 
    return ['background-color: {}'.format(color) for _ in col]
#-------------------------------------------------------------------
#----------------------PAGE ENVIRONMENT-----------------------------
#-------------------------------------------------------------------
#----PAGE_CONFIG----
st.set_page_config(page_title="KPI Dashboard WebApp",
                   page_icon=":bar_chart:",
                   layout="wide"
)
#----USER_AUTHENTICATION----
names = ["Admin"]
usernames = ["Admin"]
#----LOAD_HASHEDPASSWORD----
file_path = Path(__file__).parent / "hashed_pw.pkl"
with file_path.open("rb") as file:
    hashed_passwords = pickle.load(file)
authenticator = stauth.Authenticate(names, usernames, hashed_passwords, "kpi_dashboard", "abcdef", cookie_expiry_days=1)
name, authenticator_status, username = authenticator.login("Login","main")
current_month = datetime.now().month
current_year = datetime.now().year
month_selected = 0
year_selected = 0
#----CHECK_AUTHENTICATOR----
if authenticator_status == False:
    st.error("Username/password is incorrect")
if authenticator_status == None:
    st.warning("Please enter your username and password")
if authenticator_status == True:
    st.title(f"Welcome {name} !")
    st.write("This is a KPI analysis website, helping managers track, analyze and evaluate data collected from employee work results.")
    if(current_month == 1):
        default_month = 11
        default_year = current_year - 1
    else:
        default_month = current_month - 2
        default_year = current_year
    if "select_month" not in st.session_state:
        st.session_state["select_month"] = default_month
    get_value("select_month")
    selected_month = st.sidebar.selectbox(
        "Select Month",
        [1,2,3,4,5,6,7,8,9,10,11,12],
        key= "_select_month",
        on_change=save_value,
        args=["select_month"]
    )
    save_value("select_month")
    if "select_year" not in st.session_state:
        st.session_state["select_year"] = default_year
    get_value("select_year")
    selected_year = st.sidebar.selectbox(
        "Select Year",
        list(range(default_year-2,default_year+2)),
        key= "_select_year",
        on_change=save_value,
        args=["select_year"]
    ) 
    save_value("select_year")
    excel_name = './data/' + str(selected_year) + '-' + str(selected_month) +'.xlsx'
    checkfile = os.path.isfile(excel_name)
    if checkfile == True:
        df = pd.read_excel(io=excel_name, 
            engine='openpyxl',
            sheet_name='Data',
            skiprows=3,
            usecols='B:Q',
            nrows=1000,
    )
st.sidebar.markdown("---")
authenticator.logout("Log out","sidebar")