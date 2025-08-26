import streamlit as st

def show_app_info():
    with st.expander("📖 App Documentation", expanded=False):
        st.markdown("### Overview")
        st.info(
            "MP Data Processor allows you to upload multiple Excel/CSV files, process them according "
            "to your `groups.json` configuration, and download a standardized Media Plan."
        )



        st.markdown("### Steps to Use")
        st.markdown(
            "1. **Select the agency commission** on the left sidebar.\n"
            "2. **Click 'Browse files'**, find your folder with input Excel files, select all files using Ctrl + A, and upload them.\n"
            "3. Once files are loaded, **click the 'Process Files' button** on the sidebar to start processing.\n"
            "4. A **progress bar** will appear. Processing time can range from a few seconds to several minutes depending on the number of files.\n"
            "5. When processing completes, you will see **three tabs**:\n"
            "   - **Preview:** Check a portion of the output and filter by provider using the new sidebar button.\n"
            "   - **Download:** Download the final Excel file.\n"
            "   - **Summary:** See the total number of files, rows processed, and columns detected."
        )

        st.markdown("### Calculations - How the yellow columns are calculated")
        st.success(
            """
        Rent/month = original Rent/month * 1.2  \n
        Production = Size * 5  \n
        Posting = POSTARE FURNIZOR * 1.2  \n
        Ag Comm % = agency commission entered  \n
        Total rent = Rent/month * No. of months  \n
        Agency commission = (Posting + Production + Total rent) * Ag Comm %  \n
        Advertising taxe % = 3%  \n
        Advertising taxe = ((Total rent + Posting) * Ag Comm % + Total rent + Posting) * 3%  \n
        Total Cost = Advertising taxe + Agency commission + Posting + Production + Total rent
        """
        )

        st.markdown("### Notes")
        st.info(
            "- Uses `pandas`, `openpyxl`, `streamlit`.\n"
            "- If you encounter errors, new input structures, or unexpected results, please reach out for support."
        )
