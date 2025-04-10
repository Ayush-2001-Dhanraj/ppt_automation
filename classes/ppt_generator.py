import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from custom_presentation import CustomPresentation
import plotly.graph_objects as go
from pptx.util import Inches
import matplotlib.ticker as mticker
from scipy.interpolate import make_interp_spline
import datetime as dt


class PowerPointGenerator:
    def __init__(self, data, personal_data):
        self.data = data
        self.personal_data = personal_data

    def format_number(self, value):
        """
        Convert a numeric value to a human-readable format with suffixes (K, M, B).
        """
        if value == 0:
            return 0
        elif abs(value) >= 1_000_000:
            return f"{value / 1_000_000:.1f}M"
        elif abs(value) >= 1_000:
            return f"{value / 1_000:.1f}K"
        elif isinstance(value, int):
            return value  # Return as is if the value is an integer
        else:
            return f"{value:.2f}"  # Format to 2 decimal places if the value is a float

    def plot_time_series(self, data_grouped, filename):
        # Define filenames for saving
        filename_a = f"{filename}_a.png"
        filename_b = f"{filename}_b.png"

        x_values = range(len(data_grouped))
        # Create and save the bar plot
        plt.figure(figsize=(20, 6))  # Twice as wide
        bars = plt.bar(x_values, data_grouped["Score"], color="green")
        plt.ylabel("Runs Scored")
        plt.title("Centuries Scored Over the Years (Bar Chart)")
        plt.xlim(-0.5, len(data_grouped) - 0.5)  # Remove extra padding
        plt.gca().yaxis.set_major_locator(mticker.MaxNLocator(integer=True))
        plt.xticks([])  # Remove x-axis labels\
        plt.tight_layout()  # Optimize spacing

        for i, (bar, year, score) in enumerate(
            zip(bars, data_grouped["date"], data_grouped["Score"])
        ):
            # Annotate score on top of the bar
            plt.text(
                bar.get_x() + bar.get_width() / 2,
                bar.get_height() + 0.5,
                str(score),
                ha="center",
                va="bottom",
                fontsize=12,
                fontweight="bold",
                color="green",
            )

            # Annotate year inside the bar (rotated, at bottom)
            plt.text(
                bar.get_x() + bar.get_width() / 2,
                bar.get_height() * 0.1,
                str(year),
                ha="center",
                va="bottom",
                fontsize=12,
                color="white",
                rotation=90,
            )

        plt.figure(figsize=(20, 6))  # Twice as wide
        y_values = data_grouped["Score"]

        # Scatter points
        plt.scatter(
            x_values, y_values, color="blue", label="Scores", s=100, edgecolors="black"
        )

        # Smooth trend line (LOESS-style)
        if len(x_values) > 2:
            x_smooth = np.linspace(
                min(x_values), max(x_values), 300
            )  # More points for smoothness
            y_smooth = make_interp_spline(x_values, y_values, k=3)(
                x_smooth
            )  # Cubic spline interpolation
            plt.plot(
                x_smooth,
                y_smooth,
                color="red",
                linestyle="--",
                linewidth=2,
                label="Trend Line",
            )

        plt.ylabel("Total Centuries")
        plt.title("Total Centuries Per Year (Scatter + Trend)")
        plt.xticks([])  # Remove x-axis labels
        plt.xlim(-0.5, len(data_grouped) - 0.5)  # Remove extra padding
        plt.gca().yaxis.set_major_locator(
            mticker.MaxNLocator(integer=True)
        )  # Ensure whole numbers
        plt.legend()
        plt.tight_layout()

        # Ensure directory exists before saving
        os.makedirs(os.path.dirname(filename_b), exist_ok=True)
        plt.savefig(filename_b)
        plt.close()

        return filename_a, filename_b

    def create_graph_html_from_scores(
        self, data_grouped, slideName, key, combined_file_name
    ):
        # Prepare data
        x = list(range(len(data_grouped)))
        dates = data_grouped["date"].tolist()
        scores = data_grouped["Score"].tolist()

        # First Plot: Bar chart with annotations (Score + Year)
        fig1 = go.Figure()

        # Bar trace
        fig1.add_trace(
            go.Bar(
                x=x,
                y=scores,
                text=[str(score) for score in scores],
                textposition="outside",
                name="Scores",
                marker=dict(color="green"),
                hovertemplate="<b>%{customdata}</b><br>Score: %{y}<extra></extra>",
                customdata=dates,
            )
        )

        # Annotations for years inside bars (as shapes/text)
        for i, (xi, date, score) in enumerate(zip(x, dates, scores)):
            fig1.add_annotation(
                x=xi,
                y=score * 0.1,
                text=str(date),
                showarrow=False,
                font=dict(size=12, color="white"),
                textangle=90,
                xanchor="center",
                yanchor="bottom",
            )

        fig1.update_layout(
            title="Centuries Scored Over the Years (Bar Chart)",
            xaxis=dict(title="", showticklabels=False),
            yaxis=dict(title="Runs Scored", tickmode="linear", dtick=1),
            bargap=0.2,
            margin=dict(t=60, b=30),
            height=500,
        )

        # Second Plot: Scatter + smooth trend line
        fig2 = go.Figure()

        # Scatter trace
        fig2.add_trace(
            go.Scatter(
                x=x,
                y=scores,
                mode="markers",
                marker=dict(size=10, color="blue", line=dict(width=1, color="black")),
                name="Scores",
                hovertemplate="<b>%{customdata}</b><br>Score: %{y}<extra></extra>",
                customdata=dates,
            )
        )

        # LOESS-style trend line using cubic spline interpolation (like make_interp_spline)
        if len(x) > 2:
            from scipy.interpolate import make_interp_spline

            x_smooth = np.linspace(min(x), max(x), 300)
            y_smooth = make_interp_spline(x, scores, k=3)(x_smooth)
            fig2.add_trace(
                go.Scatter(
                    x=x_smooth,
                    y=y_smooth,
                    mode="lines",
                    name="Trend Line",
                    line=dict(color="red", dash="dash"),
                )
            )

        fig2.update_layout(
            title="Total Centuries Per Year (Scatter + Trend)",
            xaxis=dict(title="", showticklabels=False),
            yaxis=dict(title="Total Centuries", tickmode="linear", dtick=1),
            height=500,
            margin=dict(t=60, b=30),
            legend=dict(
                orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1
            ),
        )

        # Ensure directory exists before saving
        self._create_directory_for_file(combined_file_name)

        now = dt.datetime.now()
        # Write HTML
        with open(combined_file_name, "w") as f:
            f.write(f"""
            <div style="
                position: sticky; 
                top: 0; 
                background-color: #333; 
                color: white; 
                text-align: center; 
                padding: 10px; 
                font-size: 18px; 
                font-weight: bold; 
                z-index: 1000;
            ">
                {key} - {slideName} - {now.date()} 
            </div>
            """)

            f.write(fig1.to_html(full_html=False, include_plotlyjs="cdn"))
            f.write("<br><br>")
            f.write(fig2.to_html(full_html=False, include_plotlyjs=False))

            f.write(f"""
            <div style="
                background-color: #333; 
                color: white; 
                text-align: center; 
                padding: 10px; 
                font-size: 14px; 
                position: fixed; 
                bottom: 0; 
                width: 100%;
                z-index: 1000;
            ">
                Generated on: {now.date()}  | Ayush Dhanraj
            </div>
            """)

    def _create_directory_for_file(self, filename):
        dir_name = os.path.dirname(filename)
        if not os.path.exists(dir_name) and dir_name:
            os.makedirs(dir_name)

    def get_top_warranty_sub_drivers(
        self, df, start_year, start_month, end_year, end_month
    ):
        # Filter based on the given year and month range
        df_filtered = df[
            (
                (df["close_year"] > start_year)
                | (
                    (df["close_year"] == start_year)
                    & (df["close_month"] >= start_month)
                )
            )
            & (
                (df["close_year"] < end_year)
                | ((df["close_year"] == end_year) & (df["close_month"] <= end_month))
            )
        ]

        # Count occurrences of Warranty_sub_Driver
        top_warranty_sub_drivers = (
            df_filtered["Warranty_sub_Driver"].value_counts().head(4)  # Get top 4
        )

        return top_warranty_sub_drivers

    def _handle_general_flow(self, key, name):
        prs = CustomPresentation(key, name)

        total_issues = []

        filtered_Data = self.data[self.data["gender"] == name]
        players = filtered_Data["name"].unique()

        for player in players:
            # product represents one slide of the PPT
            slide_name = player
            print("Slide Name: " + slide_name)

            # Filter Calls Data
            player_df = filtered_Data[filtered_Data["name"] == player]

            # Ensure date column is in datetime format and extract the year
            player_df = player_df.copy()
            player_df.loc[:, "date"] = pd.to_datetime(
                player_df["Date"], errors="coerce"
            ).dt.year
            player_df = player_df.sort_values(by="date").reset_index(drop=True)
            player_info = self.personal_data[self.personal_data["Name"] == slide_name]

            filename_a, filename_b = self.plot_time_series(
                player_df,
                filename=f"graphs/{key}/{slide_name}/{slide_name}",
            )

            self.create_graph_html_from_scores(
                player_df,
                slide_name,
                key,
                f"PPTS/assets/{key}/{slide_name}.html",
            )

            prs.add_player_info(slide_name, player_info, total_cen=len(player_df))

            prs.add_slide(
                slide_name,
                [
                    filename_a,
                    filename_b,
                ],
                player_df,
                f"assets/{key}/{slide_name}.html",
            )

            total_issues.append(
                [slide_name, player_info["Country"].unique()[0], len(player_df)]
            )

            plt.close()

        columns = ["Name", "Country", "Total Centuries"]

        prs.add_aggregate_slide(
            key,
            pd.DataFrame(total_issues, columns=columns),
        )

        ppt_name = f"PPTS/player_{key}.pptx"
        self._create_directory_for_file(ppt_name)
        prs.save(ppt_name)
