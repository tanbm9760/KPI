from Welcome import *
#---------------------------------------------------------
st.set_page_config(page_title="KPI Dashboard WebApp",
                   page_icon=":bar_chart:",
                   layout="wide"
)
#---------------------------------------------------------
def save_value(key):
    st.session_state[key] = st.session_state["_"+key]
def get_value(key):
    st.session_state["_"+key] = st.session_state[key] 
#---------------------------------------------------------

if "select_month" not in st.session_state:
    st.session_state["select_month"] = default_month
get_value("select_month")
selected_month = st.sidebar.selectbox(
        "Select Month",
        [1,2,3,4,5,6,7,8,9,10,11,12],
        key= "_select_month",
        args=["select_month"],
        on_change=save_value,
    )
save_value("select_month")
if "select_year" not in st.session_state:
    st.session_state["select_year"] = default_year
get_value("select_year")
selected_year = st.sidebar.selectbox(
        "Select Year",
        list(range(default_year-2,default_year+2)),
        key= "_select_year",
        args=["select_year"],
        on_change=save_value,
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
#---------------------------------------------------------
with stylable_container(
        key="title",
        css_styles="""
        {   
            text-align: center;
        }"""
    ):
    st.header("3C .Inc Key Performance Indicators Dashboard - " + str(selected_month) + "/" + str(selected_year)) 
st.markdown("---")
#---------------------------------------------------------
if "select_department" not in st.session_state:
    st.session_state["select_department"] = "KD1"
get_value("select_department")
department = st.sidebar.multiselect(
    "Select Department",
    options=df["Department"].unique(),
    max_selections=3,
    key= "_select_department",
    args=["select_department"],
    on_change=save_value,
)
df_selection = df.query(
    "Department == @department"
)
#------------------------------------------------------------------------------
if department:
    st.subheader("1️⃣ Statistical Index")
    st.markdown("###")
    statistical_index(df_selection)
    st.markdown("---")
    #------------------------------------------------------------------------------
    st.subheader("2️⃣ Works Histogram")
    st.markdown("###")
    fig_11, fig_12 = st.columns([2,1])
    with fig_11:
        works_bar_sum(df_selection, "Name", "Department", "Assigned", "Assigned Works", "rgb(220,220,220)",0.8)
        works_bar_sum(df_selection, "Name", "Department", "Urgent", "Urgent Works", "rgb(38,151,215)",0.8)
        works_bar_sum(df_selection, "Name", "Department", "Important", "Important Works", "rgb(255,127,14)",0.8)
        works_bar_sum(df_selection, "Name", "Department", "Completed", "Completed Works", "rgb(44,160,44)",0.8)
        works_bar_sum(df_selection, "Name", "Department", "Late", "Completed Late Works", "rgb(214,39,40)",0.8)
        works_bar_sum(df_selection, "Name", "Department", "Out of date", "Out of date Works", "rgb(148,103,189)",0.8)
    with fig_12:
        works_bar_mean(df_selection, "Department", "Department", "Assigned", "Average Assigned Works", "rgb(220,220,220)",0.4)
        works_bar_mean(df_selection, "Department", "Department", "Urgent", "Average Urgent Works", "rgb(38,151,215)",0.4)
        works_bar_mean(df_selection, "Department", "Department", "Important", "Average Important Works", "rgb(255,127,14)",0.4)
        works_bar_mean(df_selection, "Department", "Department", "Completed", "Average Completed Works", "rgb(44,160,44)",0.4)
        works_bar_mean(df_selection, "Department", "Department", "Late", "Average Completed Late Works", "rgb(214,39,40)",0.4)
        works_bar_mean(df_selection, "Department", "Department", "Out of date", "Average Out of date Works", "rgb(148,103,189)",0.4)   
    st.markdown("---")
    #------------------------------------------------------------------------------
    st.subheader("3️⃣ Work LeaderBoards")
    st.markdown("###")
    Works_LeaderBoards(df_selection)
    st.markdown("---")
    #---------------------------------------------------------
    st.subheader("4️⃣ KPI Histogram")
    st.markdown("###")
    fig_21, fig_22 = st.columns(2)
    with fig_21:
        fig_kpifinal_his(df_selection)
    with fig_22:
        fig_coef_his(df_selection)
    st.markdown("---")
    #------------------------------------------------------------------------------
    st.subheader("5️⃣ KPI LeaderBoards")
    st.markdown("###")
    KPI_Leaderboards(df_selection)
    st.markdown("---")
    #---------------------------------------------------------
    st.subheader("6️⃣ Data Frame")
    st.markdown("###")
    st.dataframe(
        df_selection,
        width=1800,
        hide_index=True,
        column_order=["Department","Name","Assigned","Urgent","Important","Completed","Late","Out of date","Kpi Works","Kpi Bonus","Kpi Final","Coefficient","Note"],
        column_config={
          "Kpi Works": st.column_config.ProgressColumn("KPI Works", format="%.2f", min_value=50, max_value=150),
          "Kpi Bonus": st.column_config.ProgressColumn("KPI Bonus", format="%.1f", min_value=-20, max_value=20),
          "Kpi Final": st.column_config.ProgressColumn("KPI Final", format="%.2f", min_value=50, max_value=150),
          "Coefficient": st.column_config.ProgressColumn("KPI Coef", format="%.1f", min_value=0.0, max_value=2.0),
          "Note": st.column_config.TextColumn("Note", width="large"),
     }
    )
else:
    st.write("Please Select Department!")
st.sidebar.markdown("---")
authenticator.logout("Log out","sidebar")