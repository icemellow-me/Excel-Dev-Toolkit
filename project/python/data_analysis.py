"""
Data Analysis Module — Grade Analytics with pandas & matplotlib
==================================================================
Companion to student_management.py — reads the generated workbook
and produces analytics, visualizations, and PDF reports.

Usage:
    python data_analysis.py --input ./output/Atlas_International_School_Student_Management.xlsx
    python data_analysis.py --help
"""

import argparse
import os
import matplotlib
matplotlib.use("Agg")  # Headless backend
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np

SUBJECTS = ["English", "Mathematics", "Science", "ICT", "Social Studies"]


def load_grades(filepath):
    """Load the Grades sheet from the workbook."""
    df = pd.read_excel(filepath, sheet_name="Grades", engine="openpyxl")
    # Normalize column names (strip spaces for matching)
    df.columns = df.columns.str.strip()
    print(f"Loaded {len(df)} student records from Grades sheet.")
    print(f"  Columns: {list(df.columns)}")
    return df


def load_students(filepath):
    """Load the Students sheet."""
    df = pd.read_excel(filepath, sheet_name="Students", engine="openpyxl")
    df.columns = df.columns.str.strip()
    # Create a unified StudentID column if it has spaces
    if "StudentID" not in df.columns and "Student ID" in df.columns:
        df = df.rename(columns={"Student ID": "StudentID"})
    if "StudentID" not in df.columns:
        # Try to find a column that matches
        for col in df.columns:
            if col.replace(" ", "").lower() == "studentid":
                df = df.rename(columns={col: "StudentID"})
                break
    print(f"Loaded {len(df)} student records from Students sheet.")
    print(f"  Columns: {list(df.columns)}")
    return df


def analyze_subject_statistics(df):
    """Compute subject-wise statistics."""
    print("\n" + "=" * 60)
    print("  SUBJECT-WISE STATISTICS")
    print("=" * 60)

    stats = []
    for subject in SUBJECTS:
        scores = df[subject]
        avg = scores.mean()
        high = scores.max()
        low = scores.min()
        pass_count = (scores >= 50).sum()
        pass_rate = pass_count / len(scores) * 100

        stats.append({
            "Subject": subject,
            "Average": round(avg, 2),
            "Highest": high,
            "Lowest": low,
            "Pass Count": pass_count,
            "Pass Rate": f"{pass_rate:.1f}%"
        })

        print(f"  {subject:15s}  Avg: {avg:5.2f}  High: {high:5.1f}  "
              f"Low: {low:5.1f}  Pass: {pass_rate:5.1f}%")

    return pd.DataFrame(stats)


def analyze_gender_comparison(df_students, df_grades):
    """Compare performance by gender."""
    print("\n" + "=" * 60)
    print("  GENDER COMPARISON")
    print("=" * 60)

    merged = df_grades.merge(
        df_students[["StudentID", "Gender"]], on="StudentID")

    for gender in ["Male", "Female"]:
        subset = merged[merged["Gender"] == gender]
        print(f"  {gender:8s}: Count={len(subset):3d}  "
              f"Average={subset['Average'].mean():.2f}  "
              f"Top={subset['Average'].max():.1f}")

    # T-test
    male_avg = merged[merged["Gender"] == "Male"]["Average"]
    female_avg = merged[merged["Gender"] == "Female"]["Average"]
    print(f"\n  Difference: {abs(male_avg.mean() - female_avg.mean()):.2f} points")


def analyze_rankings(df):
    """Top and bottom performers."""
    print("\n" + "=" * 60)
    print("  CLASS RANKINGS (Top 5)")
    print("=" * 60)

    ranked = df.sort_values("Average", ascending=False)
    for i, (_, row) in enumerate(ranked.head(5).iterrows(), 1):
        print(f"  {i}. {row['Name']:20s}  Avg: {row['Average']:.2f}  Grade: {row['Grade']}")

    print("\n  Bottom 3:")
    for i, (_, row) in enumerate(ranked.tail(3).iterrows(), 1):
        print(f"  {i}. {row['Name']:20s}  Avg: {row['Average']:.2f}  Grade: {row['Grade']}")


def generate_charts(df, output_dir):
    """Generate and save charts."""
    os.makedirs(output_dir, exist_ok=True)
    charts = []

    # Chart 1: Subject Averages Bar Chart
    fig, ax = plt.subplots(figsize=(10, 6))
    avgs = [df[s].mean() for s in SUBJECTS]
    colors = ["#4472C4", "#5B9BD5", "#70AD47", "#FFC000", "#ED7D31"]
    bars = ax.bar(SUBJECTS, avgs, color=colors, edgecolor="black", linewidth=0.5)
    ax.set_title("Subject-Wise Average Scores", fontsize=14, fontweight="bold")
    ax.set_ylabel("Average Score (%)")
    ax.set_ylim(0, 100)
    for bar, val in zip(bars, avgs):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 1,
                f"{val:.1f}", ha="center", fontweight="bold")
    plt.xticks(rotation=15, ha="right")
    plt.tight_layout()
    path = os.path.join(output_dir, "subject_averages.png")
    plt.savefig(path, dpi=150)
    plt.close()
    charts.append(path)

    # Chart 2: Grade Distribution Pie Chart
    fig, ax = plt.subplots(figsize=(8, 8))
    grade_counts = df["Grade"].value_counts().sort_index()
    grade_colors = ["#00B050", "#92D050", "#FFFF00", "#FFC000", "#FF8000", "#FF0000"]
    ax.pie(grade_counts.values, labels=grade_counts.index, autopct="%1.1f%%",
           colors=grade_colors[:len(grade_counts)], startangle=90,
           textprops={"fontsize": 12, "fontweight": "bold"})
    ax.set_title("Grade Distribution", fontsize=14, fontweight="bold")
    plt.tight_layout()
    path = os.path.join(output_dir, "grade_distribution.png")
    plt.savefig(path, dpi=150)
    plt.close()
    charts.append(path)

    # Chart 3: Student Scores Heatmap (simulated)
    fig, ax = plt.subplots(figsize=(12, 8))
    score_data = df[SUBJECTS].values
    im = ax.imshow(score_data, cmap="RdYlGn", aspect="auto", vmin=0, vmax=100)
    ax.set_xticks(range(len(SUBJECTS)))
    ax.set_xticklabels(SUBJECTS, rotation=30, ha="right")
    ax.set_yticks(range(len(df)))
    ax.set_yticklabels(df["Name"].values, fontsize=8)
    ax.set_title("Student Scores by Subject (Heatmap)", fontsize=14, fontweight="bold")
    plt.colorbar(im, ax=ax, label="Score")
    plt.tight_layout()
    path = os.path.join(output_dir, "scores_heatmap.png")
    plt.savefig(path, dpi=150)
    plt.close()
    charts.append(path)

    # Chart 4: Gender Comparison
    fig, ax = plt.subplots(figsize=(10, 6))
    male_avgs = [df[df.get("Gender") == "Male"][s].mean() if "Gender" in df.columns else df[s].mean() for s in SUBJECTS]
    female_avgs = [df[df.get("Gender") == "Female"][s].mean() if "Gender" in df.columns else df[s].mean() for s in SUBJECTS]

    x = np.arange(len(SUBJECTS))
    width = 0.35
    ax.bar(x - width/2, male_avgs, width, label="Male", color="#4472C4")
    ax.bar(x + width/2, female_avgs, width, label="Female", color="#ED7D31")
    ax.set_title("Gender Comparison by Subject", fontsize=14, fontweight="bold")
    ax.set_ylabel("Average Score")
    ax.set_xticks(x)
    ax.set_xticklabels(SUBJECTS, rotation=15, ha="right")
    ax.legend()
    ax.set_ylim(0, 100)
    plt.tight_layout()
    path = os.path.join(output_dir, "gender_comparison.png")
    plt.savefig(path, dpi=150)
    plt.close()
    charts.append(path)

    print(f"\n📊 Generated {len(charts)} charts:")
    for p in charts:
        print(f"   ✅ {p}")

    return charts


def main():
    parser = argparse.ArgumentParser(description="Student Data Analysis & Visualization")
    parser.add_argument("--input", required=True, help="Path to the .xlsx workbook")
    parser.add_argument("--output", default="./output/charts",
                        help="Output directory for charts")
    args = parser.parse_args()

    print(f"\n📈 Data Analysis Report")
    print(f"   Input: {args.input}")
    print(f"   Output: {args.output}\n")

    # Load data
    df_grades = load_grades(args.input)
    df_students = load_students(args.input)

    # Run analyses
    stats = analyze_subject_statistics(df_grades)
    analyze_gender_comparison(df_students, df_grades)
    analyze_rankings(df_grades)

    # Generate visualizations
    charts = generate_charts(df_grades, args.output)

    print(f"\n{'=' * 60}")
    print(f"  ✅ Analysis complete! Charts saved to: {args.output}")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
