import pandas as pd
import numpy as np
import pandas as pd
import numpy as np
from scipy.stats import linregress

class QPCRAnalysis:
    "class to caclulate RSQ ,effciencey , slope and intercept"
    def __init__(self, df):
        """
        Initializes the class with a DataFrame containing 'Log template Conc.' and 'CT' values.
        """
        self.df = df
        self.slope = None
        self.intercept = None
        self.rsq = None
        self.efficiency = None
    
    def calculate_rsq_efficiency(self):
        """
        Calculates the R-squared (RSQ) and Efficiency for the qPCR assay based on dilution and Ct values
        """
        # Perform linear regression: Ct = slope * log_dilution + intercept
        self.slope, self.intercept, r_value, p_value, std_err = linregress(self.df['Log template Conc.'], self.df['CT'])

        # Calculate R-squared (RSQ)
        self.rsq = r_value ** 2

        # Calculate Efficiency from the slope
        self.efficiency = -1 + 10**(-1/self.slope)

        # Print slope, intercept, and raw efficiency value
        print(f"Slope (rounded to 3 decimal places): {self.slope}")
        print(f"Intercept: {self.intercept:.8f}")
        print(f"Raw Efficiency Value: {self.efficiency:.8f}")

    def get_slope_intercept(self):
        return self.slope, self.intercept 
    
def calculate_vg_rx(df, intercept, slope):
    # function to caclulate rx well and /ml values for VG used in sample analysis and summary stats
    def calculate_dilution(ct):
        if pd.isna(ct) or ct == "":
            return None
        return 10 ** ((ct - intercept) / slope)
    
    # Apply the function only to rows where 'Sample Type' is "Sample", else set None
    df["VG/rx well"] = np.where(
        df["Sample Type"] == "Sample",
        df["Ct/Cq"].apply(calculate_dilution),
        None
    )
    # Calculate VG/ml based on VG/rx well and Dilution, with formula E3*1000/10*B3*10
    df["VG/ml"] = np.where(
        (df["Sample Type"] == "Sample") & (df["VG/rx well"].notna()),
        df["VG/rx well"].astype(float) * 1000 / 10 * df["Dilution"] * 10,
        None
    )
    # Filter the DataFrame to include only rows where Sample Type is "Sample"
    df_filtered = df[df["Sample Type"] == "Sample"].copy()
    df_filtered = df_filtered.sort_values(by=["Sample Name", "Dilution"], ascending=[True, True])
    return df_filtered







