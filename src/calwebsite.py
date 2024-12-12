import streamlit as st


def calculate(num1, num2, operation):
    """Perform calculation based on the selected operation."""
    try:
        if operation == 'Add':
            return num1 + num2
        elif operation == 'Subtract':
            return num1 - num2
        elif operation == 'Multiply':
            return num1 * num2
        elif operation == 'Divide':
            if num2 == 0:
                raise ValueError("Cannot divide by zero")
            return num1 / num2
    except Exception as e:
        st.error(f"Calculation error: {e}")
        return None


def main():
    # Configure page
    st.set_page_config(page_title="Simple Calculator", page_icon="🧮")

    # Title
    st.title("Simple Calculator")

    # Create input columns
    col1, col2 = st.columns(2)

    # Number inputs
    with col1:
        num1 = st.number_input("First Number", value=0.0, format="%.2f")

    with col2:
        num2 = st.number_input("Second Number", value=0.0, format="%.2f")

    # Operation selection
    operation = st.selectbox(
        "Select Operation",
        ["Add", "Subtract", "Multiply", "Divide"]
    )

    # Calculate button
    if st.button("Calculate"):
        result = calculate(num1, num2, operation)

        if result is not None:
            st.success(f"Result of {operation}: {result:.2f}")


# Ensure the app runs only when called directly
if __name__ == "__main__":
    main()