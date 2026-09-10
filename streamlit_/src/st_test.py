import streamlit as st
import subprocess
from constants import pthPy, prp
from utilities import timediff


st.title("Proportional Portfolio Valuation Report")

st.write("Do the proportional valuation!")

if st.button("Click em to do the proportional valuation!"):
    subprocess.run(["python", prp])
