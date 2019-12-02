# Global variables
line_color = '#3f51b5'      # line (trace0) color for any plot
marker_color = '#43a047'    # marker color for any plot
marker_border_color = '#ffffff'     # marker border color for any plot
cl_color = '#ffa000'    # control limit line color for any plot
sl_color = '#e53935'    # spec limit line color for any plot

# NIT ER Plot
er_nit_plot_title = 'Nit ER Plot for REPL1A'  # title for Nit ER plot
er_nit_plot_xlabel = 'Date'        # xaxis name for Nit ER plot
er_nit_plot_ylabel = 'Nit ER (A/min)'   # yaxis name for Nit ER plot
er_nit_plot_html_file = 'REPL1A_Nit_ER-Plot.html'   # HTML filename for Nit ER plot
er_nit_plot_trace_count = 5    # no. of traces in Nit ER plot

# Sheet names
sht_name_er_nit = 'REPL1A-ERNit'

# Columns
sht_er_nit_columns = ["Date (MM/DD/YYYY)", "Etch Rate (A/Min)", "% Uni", "LSL", "USL", "LCL", "UCL", "Remarks", "% Uni USL", "% Uni UCL"]

# Date formatter
date_format = "%m-%d-%Y %H:%M:%S"

"""Skiprows"""
skiprows_nit = 10

excel_file_directory = "1.xlsx"