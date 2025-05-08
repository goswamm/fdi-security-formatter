import streamlit as st
import pandas as pd
import xlsxwriter
import io

# Mapping of Security Object Code to Fusion Security Context Name
security_context_map = {
    "OA4F_SEC_FIN_AP_BUSINESSUNIT_LIST": "AP Business Units",
    "OA4F_SEC_FIN_AR_BUSINESSUNIT_LIST": "AR Business Units",
    "OA4F_SEC_CST_COST_ORG_LIST": "Cost Organizations",
    "OA4F_SEC_FIN_FA_ASSET_BOOK_LIST": "FA Asset Book",
    "OA4F_SEC_HCM_BUSINESSUNIT_LIST": "HCM Business Units",
    "OA4F_SEC_HCM_COUNTRY_LIST": "HCM Country List",
    "OA4F_SEC_HCM_DEPARTMENT_LIST": "HCM Departments",
    "OA4F_SEC_HCM_LEGAL_EMPLOYER_LIST": "HCM Legal Employers",
    "OA4F_SEC_HCM_SEE_SELF_RECORD": "HCM Show Self Record",
    "OA4F_SEC_INV_BUSINESSUNIT_LIST": "Inventory Business Units",
    "OA4F_SEC_INV_ORG_TRANSACTIONS_LIST": "Inventory Organizations",
    "OA4F_SEC_FIN_LEDGER_LIST": "Ledgers",
    "OA4F_SEC_OM_BUSINESS_UNIT_LIST": "Order Management Business Units",
    "OA4F_SEC_PPM_PROJECT_BUSINESSUNIT_LIST": "Project Business Units",
    "OA4F_SEC_PPM_EXPENDITURE_BUSINESSUNIT_LIST": "Project Expenditure Business Units",
    "OA4F_SEC_PPM_PROJECT_ORGANIZATION_LIST": "Project Organizations",
    "OA4F_SEC_PROC_REQ_BUSINESSUNIT_LIST": "Requisition Business Units",
    "OA4F_SEC_PROC_SPEND_PRC_BUSINESSUNIT_LIST": "Spend Procurement Business Units"
    # Add the rest here...
}

def generate_excel(df):
    df = df.rename(columns={
        "USERNAME": "User",
        "SEC_OBJ_CODE": "Security Object Code",
        "SEC_OBJ_MEMBER_VAL": "Member Value",
        "SEC_OBJ_MEMBER_NAME": "Member Name",
        "OPERATION": "Operation"
    })
    df["Assignment"] = "*"
    
    pivot_df = df.pivot_table(
        index="User",
        columns=["Security Object Code", "Member Name"],
        values="Assignment",
        aggfunc="first"
    ).reset_index()

    user_col = pivot_df["User"]
    data_only = pivot_df.drop(columns="User")
    sorted_columns = sorted(data_only.columns, key=lambda x: (x[0], x[1]))
    data_only = data_only[sorted_columns]
    pivot_df_sorted = pd.concat([user_col, data_only], axis=1)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        pivot_df_sorted.to_excel(writer, index=False, sheet_name='Assignments', startrow=3, header=False)
        workbook = writer.book
        worksheet = writer.sheets['Assignments']

        header_format = workbook.add_format({
            'text_wrap': True, 'rotation': 90, 'align': 'center',
            'valign': 'bottom', 'border': 1
        })
        x_format = workbook.add_format({
            'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#BDD7EE'
        })
        fill_colors = ['#D9E1F2', '#FCE4D6']

        worksheet.write(0, 0, "Fusion Security Context Name", header_format)
        worksheet.write(1, 0, "Security Object Code", header_format)
        worksheet.write(2, 0, "User", header_format)

        col_start = 1
        prev_code = None
        color_index = 0

        while col_start <= len(sorted_columns):
            sec_code = sorted_columns[col_start - 1][0]
            mem_names = []
            col_end = col_start

            while col_end <= len(sorted_columns) and sorted_columns[col_end - 1][0] == sec_code:
                mem_names.append(sorted_columns[col_end - 1][1])
                col_end += 1

            color_index = 1 - color_index
            fill_format = workbook.add_format({
                'text_wrap': True, 'rotation': 90, 'align': 'center',
                'valign': 'bottom', 'border': 1, 'bg_color': fill_colors[color_index]
            })

            fusion_name = security_context_map.get(sec_code, sec_code)
            worksheet.merge_range(0, col_start, 0, col_end - 1, fusion_name, fill_format)
            worksheet.merge_range(1, col_start, 1, col_end - 1, sec_code, fill_format)

            for idx, col in enumerate(range(col_start, col_end)):
                worksheet.write(2, col, mem_names[idx], fill_format)

            col_start = col_end

        for row in range(3, 3 + len(pivot_df_sorted)):
            for col in range(1, len(pivot_df_sorted.columns)):
                value = pivot_df_sorted.iloc[row - 3, col]
                if value == "*":
                    worksheet.write(row, col, "X", x_format)

        worksheet.set_column(0, 0, 25)
        for col in range(1, len(sorted_columns) + 1):
            worksheet.set_column(col, col, 4)

    output.seek(0)
    return output

# Streamlit UI
st.title("Fusion Security Assignment Formatter")
st.write("Upload your security CSV file to generate a formatted Excel report.")

uploaded_file = st.file_uploader("Upload CSV", type=["csv"])

if uploaded_file is not None:
    df = pd.read_csv(uploaded_file)
    st.success("File uploaded successfully!")

    if st.button("Generate Excel Report"):
        excel_file = generate_excel(df)
        st.download_button(
            label="Download Formatted Excel",
            data=excel_file,
            file_name="Formatted_Security_Assignment.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
