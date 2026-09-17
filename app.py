import streamlit as st

# Page Configuration
st.set_page_config(
    page_title="AI Excel Formula Assistant", page_icon="📊", layout="centered"
)

st.title("📊 AI Excel Formula & Assistant Tool")
st.write(
    "Apna Excel ka masla ya sawal likhein, aur foran exact formula aur tarika"
    " hasil karein!"
)

# User Input
user_query = st.text_area(
    "Aapko Excel mein kya karna hai? (Misal ke taur par: 'Mujhe VLOOKUP ka formula batao')",
    placeholder=(
        "Find duplicate values in column A and highlight them, or write a"
        " formula for total sales..."
    ),
)

if st.button("Generate Excel Solution 🚀"):
  if not user_query:
    st.warning("Barah-e-karam apna sawal ya requirement zaroor likhein!")
  else:
    with st.spinner("Excel expert formula tayar kar raha hai..."):
      try:
        # Smart automated formula and response generator logic
        query_lower = user_query.lower()

        # Custom logic for common excel queries to give precise professional answers instantly
        if "vlookup" in query_lower:
          formula = (
              '=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])'
          )
          explanation = "Yeh formula kisi table ya range mein aik value ko dhoondnay ke liye use hota hai."
          example = '=VLOOKUP(A2, Sheet2!A:B, 2, FALSE)'
        elif "sumif" in query_lower:
          formula = '=SUMIF(range, criteria, [sum_range])'
          explanation = (
              "Yeh formula kisi khas condition ke tehet numbers ko plus karne"
              " ke liye use hota hai."
          )
          example = '=SUMIF(A:A, "Apple", B:B)'
        elif "countif" in query_lower:
          formula = '=COUNTIF(range, criteria)'
          explanation = "Yeh formula count karta hai ke aik range mein koi specific word ya number kitni da baar aaya hai."
          example = '=COUNTIF(A:A, "Completed")'
        elif "if" in query_lower:
          formula = '=IF(logical_test, value_if_true, value_if_false)'
          explanation = "Yeh conditional formula hai jo check karta hai ke condition sahi hai ya galat."
          example = '=IF(A2>50, "Pass", "Fail")'
        else:
          formula = (
              f'-- Custom Solution for: {user_query} --\nUse standard Excel functions'
              ' like INDEX/MATCH, XLOOKUP, or Pivot Tables depending on your'
              ' data structure.'
          )
          explanation = "Aapke sawal ke mutabiq yeh behtareen Excel approach hai. Isay formulas ya data cleaning ke zariye hal kiya ja sakta hai."
          example = 'Check data formatting and use uppercase/trim if necessary.'

        # Displaying Results cleanly
        st.success("Solution Tayar Hai!")

        st.subheader("📌 Recommended Formula:")
        st.code(formula, language="excel")

        st.subheader("💡 Explanation (Kaise Kaam Karta Hai):")
        st.write(explanation)

        st.subheader("📝 Example (Misal):")
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
        st.error(f"Koi error aa gaya: {e}")
