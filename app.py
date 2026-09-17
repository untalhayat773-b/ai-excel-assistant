import streamlit as st

# Page Configuration
st.set_page_config(
    page_title="AI Excel Formula Assistant", page_icon="📊", layout="centered"
)

st.title("📊 AI Excel Formula & Assistant Tool")
st.write(
    "Type your Excel problem below, or explore the complete formula reference"
    " directory with Math, Lookup, and Financial formulas!"
)

# Tabs for better UI/UX organization
tab1, tab2 = st.tabs(["🔍 Ask AI Assistant", "📚 Formula Reference Directory"])

with tab1:
  # User Input Section
  user_query = st.text_area(
      "What do you want to do in Excel?",
      placeholder=(
          "e.g., Calculate average sales, write a VLOOKUP formula, or round"
          " numbers..."
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
                '=VLOOKUP(lookup_value, table_array, col_index_num,'
                ' [range_lookup])'
            )
            explanation = "Searches for a value in the leftmost column of a table, and returns a value in the same row from a specified column."
            example = '=VLOOKUP(A2, Sheet2!A:B, 2, FALSE)'
          elif "xlookup" in query_lower:
            formula = (
                '=XLOOKUP(lookup_value, lookup_array, return_array,'
                ' [if_not_found])'
            )
            explanation = "Searches a range or an array, and returns the item corresponding to the first match found."
            example = '=XLOOKUP(E2, A:A, B:B, "Not Found")'
          elif "sum" in query_lower and "if" not in query_lower:
            formula = '=SUM(number1, [number2], ...)'
            explanation = (
                "Adds all the numbers in a range of cells automatically."
            )
            example = "=SUM(A1:A10)"
          elif "average" in query_lower:
            formula = '=AVERAGE(number1, [number2], ...)'
            explanation = "Calculates the arithmetic mean of a given set of numbers."
            example = "=AVERAGE(B1:B20)"
          elif "round" in query_lower:
            formula = '=ROUND(number, num_digits)'
            explanation = "Rounds a number to a specified number of digits."
            example = "=ROUND(A1, 2)"
          elif "sumif" in query_lower:
            formula = '=SUMIF(range, criteria, [sum_range])'
            explanation = "Adds numbers in a range that meet a single specific condition."
            example = '=SUMIF(A:A, "Apple", B:B)'
          elif "countif" in query_lower:
            formula = '=COUNTIF(range, criteria)'
            explanation = "Counts the number of cells within a range that meet a single condition."
            example = '=COUNTIF(A:A, "Completed")'
          elif "if" in query_lower:
            formula = '=IF(logical_test, value_if_true, value_if_false)'
            explanation = "Performs a logical test and returns one value if TRUE, another if FALSE."
            example = '=IF(A2>50, "Pass", "Fail")'
          else:
            formula = (
                f'-- Custom Solution for: {user_query} --\nUse standard Excel'
                ' functions like SUM, AVERAGE, VLOOKUP, or Pivot Tables based on'
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

with tab2:
  st.subheader("📚 Complete Excel Formula & Math Directory")
  st.write(
      "Explore core formulas including Math, Lookup, Logical, and Statistical"
      " functions."
  )

  formulas_db = [
      {
          "Category": "Math & Trig",
          "Name": "SUM",
          "Question": "How to add multiple numbers or cells together?",
          "Syntax": "=SUM(number1, [number2], ...)",
      },
      {
          "Category": "Math & Trig",
          "Name": "AVERAGE",
          "Question": "How to calculate the average of a range?",
          "Syntax": "=AVERAGE(number1, [number2], ...)",
      },
      {
          "Category": "Math & Trig",
          "Name": "ROUND",
          "Question": "How to round numbers to decimal places?",
          "Syntax": "=ROUND(number, num_digits)",
      },
      {
          "Category": "Lookup & Reference",
          "Name": "VLOOKUP",
          "Question": "How to find a value vertically in a table?",
          "Syntax": (
              "=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])"
          ),
      },
      {
          "Category": "Lookup & Reference",
          "Name": "XLOOKUP",
          "Question": "How to search values flexibly in any direction?",
          "Syntax": (
              "=XLOOKUP(lookup_value, lookup_array, return_array,"
              " [if_not_found])"
          ),
      },
      {
          "Category": "Logical",
          "Name": "IF",
          "Question": "How to apply a conditional true/false check?",
          "Syntax": "=IF(logical_test, value_if_true, value_if_false)",
      },
      {
          "Category": "Math & Trig",
          "Name": "SUMIF",
          "Question": "How to sum numbers based on specific criteria?",
          "Syntax": '=SUMIF(range, criteria, [sum_range])',
      },
      {
          "Category": "Statistical",
          "Name": "COUNTIF",
          "Question": "How to count cells matching a specific condition?",
          "Syntax": '=COUNTIF(range, criteria)',
      },
  ]

  for item in formulas_db:
    with st.expander(f"🔹 {item['Name']} ({item['Category']})"):
      st.markdown(f"**Sample Question:** {item['Question']}")
      st.code(item["Syntax"], language="excel")
