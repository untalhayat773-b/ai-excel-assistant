import streamlit as st

# Page Configuration
st.set_page_config(
    page_title="AI Excel Formula Assistant", page_icon="📊", layout="centered"
)

st.title("📊 AI Excel Formula & Assistant Tool")
st.write(
    "Type your Excel problem or requirement below, and get the exact formula"
    " and professional solution instantly!"
)

# User Input
user_query = st.text_area(
    "What do you want to do in Excel?",
    placeholder=(
        "e.g., Find duplicate values in column A and highlight them, or write a"
        " formula for total sales..."
    ),
)

if st.button("Generate Excel Solution 🚀"):
  if not user_query:
    st.warning("Please enter your Excel question or requirement first!")
  else:
    with st.spinner("AI Excel expert is generating your solution..."):
      try:
        query_lower = user_query.lower()

        # Smart automated formula and response generator logic
        if "vlookup" in query_lower:
          formula = (
              '=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])'
          )
          explanation = "This formula is used to search for a value in the leftmost column of a table, and then return a value in the same row from a column you specify."
          example = '=VLOOKUP(A2, Sheet2!A:B, 2, FALSE)'
        elif "sumif" in query_lower:
          formula = '=SUMIF(range, criteria, [sum_range])'
          explanation = "This formula adds all numbers in a range of cells that meet a single specified condition or criteria."
          example = '=SUMIF(A:A, "Apple", B:B)'
        elif "countif" in query_lower:
          formula = '=COUNTIF(range, criteria)'
          explanation = "This formula counts the number of cells within a range that meet a single condition."
          example = '=COUNTIF(A:A, "Completed")'
        elif "if" in query_lower:
          formula = '=IF(logical_test, value_if_true, value_if_false)'
          explanation = "This formula checks whether a condition is met, returning one value if TRUE and another if FALSE."
          example = '=IF(A2>50, "Pass", "Fail")'
        else:
          formula = (
              f'-- Custom Solution for: {user_query} --\nUse standard Excel'
              ' functions like INDEX/MATCH, XLOOKUP, or Pivot Tables based on'
              ' your data.'
          )
          explanation = "This is the best professional approach for your specific Excel requirement. Ensure your data is structured cleanly."
          example = 'Check data formatting, remove leading spaces, and retry.'

        # Displaying Results cleanly
        st.success("Solution Generated Successfully!")

        st.subheader("📌 Recommended Formula:")
        st.code(formula, language="excel")

        st.subheader("💡 Explanation:")
        st.write(explanation)

        st.subheader("📝 Example Usage:")
        st.code(example, language="excel")

        # Download solution text feature
        solution_text = (
            f"Excel Query: {user_query}\nFormula: {formula}\nExplanation:"
            f" {explanation}\nExample: {example}"
        )
        st.download_button(
            label="📥 Download Solution as Text File",
            data=solution_text,
            file_name="excel_solution.txt",
            mime="text/plain",
        )

      except Exception as e:
        st.error(f"An error occurred: {e}")
