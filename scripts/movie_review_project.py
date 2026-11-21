"""
Movie Review Analysis - simple step-by-step project (HCL training level)
Uses only: lists, tuples, dictionaries, loops, numpy, pandas, matplotlib

How to run: open this file in VS Code and click the Run Python File button.
If movie_reviews_raw.xlsx is missing, this program creates it automatically.
All charts open in ONE window. Close it to continue.
If Excel reading fails, run once in the terminal:  pip install openpyxl
"""

# ======================================================================
# Step 1 - Import libraries
# ======================================================================
import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# Work inside the folder that contains this file
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# Show full tables in the terminal
pd.set_option("display.width", 200)
pd.set_option("display.max_columns", None)

EXCEL_FILE = "movie_reviews_raw.xlsx"

# ======================================================================
# Step 2 - Create the practice Excel file (only if it is missing)
# ======================================================================
# Each movie is a TUPLE: (title, genre, language, quality)
# "quality" is the average rating people usually give that movie.
if not os.path.exists(EXCEL_FILE):
    np.random.seed(7)

    movies = [
        ("Silent Horizon", "Drama", "English", 7.8),
        ("Golden Empire", "Action", "Hindi", 6.2),
        ("Broken Promise", "Drama", "Hindi", 7.1),
        ("Midnight River", "Thriller", "English", 6.9),
        ("Crimson Shadow", "Horror", "English", 5.4),
        ("Hidden Journey", "Comedy", "Tamil", 6.5),
        ("Eternal Garden", "Romance", "Hindi", 6.8),
        ("Wild Signal", "Sci-Fi", "English", 7.4),
        ("Lost Kingdom", "Action", "Telugu", 6.0),
        ("Rising Echo", "Drama", "Korean", 8.0),
        ("Neon Storm", "Sci-Fi", "Korean", 7.2),
        ("Secret Legacy", "Thriller", "Tamil", 6.6),
        ("Final Voyage", "Animation", "English", 8.2),
        ("Distant Whisper", "Horror", "Hindi", 5.1),
        ("Happy Harvest", "Comedy", "Hindi", 6.3),
        ("Silent Empire", "Action", "English", 5.9),
        ("Golden River", "Romance", "Tamil", 6.7),
        ("Broken Signal", "Sci-Fi", "Telugu", 6.4),
        ("Wild Garden", "Animation", "Hindi", 7.9),
        ("Lost Echo", "Thriller", "Korean", 7.0),
    ]
    reviewers = ["Aarav", "Diya", "Rohan", "Meera", "Kabir", "Sneha",
                 "Arjun", "Priya", "Vivaan", "Ananya", "Karthik", "Riya"]

    rows = []
    for i in range(1, 201):
        movie = movies[np.random.randint(0, len(movies))]
        title, genre, language, quality = movie            # tuple unpacking
        reviewer = reviewers[np.random.randint(0, len(reviewers))]
        year = np.random.randint(2018, 2026)
        month = np.random.randint(1, 13)
        day = np.random.randint(1, 28)
        review_date = f"{year}-{month:02d}-{day:02d}"
        rating = round(np.random.normal(quality, 1.0), 1)
        rating = float(np.clip(rating, 2, 10))
        likes = np.random.randint(0, 500)
        rows.append([f"R{i:03d}", title, genre, language, reviewer, review_date, rating, likes])

    raw = pd.DataFrame(rows, columns=["Review_ID", "Movie", "Genre", "Language",
                                      "Reviewer", "Review_Date", "Rating", "Likes"])

    # ---- add messy data on purpose (for cleaning practice) ----
    for i in range(0, 200, 17):                      # extra spaces + CAPITAL letters
        raw.loc[i, "Genre"] = " " + raw.loc[i, "Genre"].upper() + " "
    for i in range(3, 200, 23):                      # small letters
        raw.loc[i, "Language"] = raw.loc[i, "Language"].lower()
    for i in range(5, 200, 31):                      # missing ratings
        raw.loc[i, "Rating"] = np.nan
    raw.loc[10, "Rating"] = 0                        # invalid ratings
    raw.loc[50, "Rating"] = -1
    raw.loc[90, "Rating"] = 15
    for i in (20, 60, 100):                          # missing reviewer
        raw.loc[i, "Reviewer"] = np.nan
    raw.loc[30, "Review_Date"] = "not available"     # bad dates
    raw.loc[70, "Review_Date"] = np.nan
    raw.loc[40, "Likes"] = -5                        # invalid likes

    raw = pd.concat([raw, raw.head(5)], ignore_index=True)   # 5 duplicate rows
    raw.to_excel(EXCEL_FILE, index=False)
    print("Created practice file:", EXCEL_FILE, "with", len(raw), "rows")

# ======================================================================
# Step 3 - Load the Excel file
# ======================================================================
df = pd.read_excel(EXCEL_FILE)
raw_rows = len(df)

print("Dataset loaded successfully")
print("Rows and columns:", df.shape)
print(df.head())

# ======================================================================
# Step 4 - Understand the data
# ======================================================================
print("\nColumn names:", df.columns.tolist())
print("\nData types:")
print(df.dtypes)
print("\nMissing values in each column:")
print(df.isnull().sum())

# ======================================================================
# Step 5 - Remove duplicate rows
# ======================================================================
duplicates = df.duplicated().sum()
print("\nDuplicate rows found:", duplicates)

df = df.drop_duplicates()
print("Rows after removing duplicates:", len(df))

# ======================================================================
# Step 6 - Clean text columns (loop)
# ======================================================================
# Remove extra spaces and make the first letter of each word capital.
text_columns = ["Movie", "Genre", "Language", "Reviewer"]

for column in text_columns:
    df[column] = df[column].str.strip().str.title()

print(df[text_columns].head())

# ======================================================================
# Step 7 - Correct data types
# ======================================================================
# Wrong values (like "not available") become missing (NaT / NaN).
df["Review_Date"] = pd.to_datetime(df["Review_Date"], errors="coerce")
df["Rating"] = pd.to_numeric(df["Rating"], errors="coerce")
df["Likes"] = pd.to_numeric(df["Likes"], errors="coerce")

# ======================================================================
# Step 8 - Find and fix invalid values
# ======================================================================
# Rating must be from 1 to 10. Likes cannot be negative.
bad_rating = (df["Rating"] < 1) | (df["Rating"] > 10)
bad_likes = df["Likes"] < 0

print("Invalid ratings found:", bad_rating.sum())
print("Invalid likes found  :", bad_likes.sum())

df.loc[bad_rating, "Rating"] = np.nan
df.loc[bad_likes, "Likes"] = np.nan

# ======================================================================
# Step 9 - Handle missing values
# ======================================================================
# Rating and Review_Date are needed, so remove rows that miss them.
# Missing reviewer -> "Anonymous". Missing likes -> median value.
rows_before = len(df)

df = df.dropna(subset=["Rating", "Review_Date"])
df["Reviewer"] = df["Reviewer"].fillna("Anonymous")
df["Likes"] = df["Likes"].fillna(df["Likes"].median())

print("Rows removed:", rows_before - len(df))
print("Rows remaining:", len(df))
print(df.isnull().sum())

# ======================================================================
# Step 10 - Create new columns (loop + if / elif / else)
# ======================================================================
df["Year"] = df["Review_Date"].dt.year
df["Month"] = df["Review_Date"].dt.month

categories = []
for r in df["Rating"]:
    if r >= 8:
        categories.append("Excellent")
    elif r >= 7:
        categories.append("Good")
    elif r >= 6:
        categories.append("Average")
    else:
        categories.append("Poor")
df["Rating_Category"] = categories

print(df[["Movie", "Year", "Rating", "Rating_Category"]].head())

# ======================================================================
# Step 11 - Key numbers (numpy + pandas)
# ======================================================================
total_reviews = len(df)
total_movies = df["Movie"].nunique()
average_rating = np.mean(df["Rating"])
total_likes = df["Likes"].sum()

print("\nTotal reviews :", total_reviews)
print("Total movies  :", total_movies)
print("Average rating:", round(average_rating, 2), "/ 10")
print("Total likes   :", int(total_likes))

# ======================================================================
# Step 12 - Rating by genre, language and year (groupby)
# ======================================================================
genre_avg = df.groupby("Genre")["Rating"].mean().round(2).sort_values(ascending=False)
language_avg = df.groupby("Language")["Rating"].mean().round(2).sort_values(ascending=False)
year_avg = df.groupby("Year")["Rating"].mean().round(2)

print("\nAverage rating by genre:")
print(genre_avg)
print("\nAverage rating by language:")
print(language_avg)
print("\nAverage rating by year:")
print(year_avg)

# ======================================================================
# Step 13 - Best and worst movies (list of tuples + loop)
# ======================================================================
# Only movies with at least 5 reviews are compared.
movie_avg = df.groupby("Movie")["Rating"].agg(["count", "mean"]).round(2)
movie_avg = movie_avg[movie_avg["count"] >= 5]
movie_avg = movie_avg.sort_values("mean", ascending=False)

best_5 = list(zip(movie_avg.head(5).index, movie_avg.head(5)["mean"]))
worst_5 = list(zip(movie_avg.tail(5).index, movie_avg.tail(5)["mean"]))

print("\nTop 5 movies:")
rank = 1
for name, avg in best_5:             # tuple unpacking
    print(rank, name, "-", avg)
    rank = rank + 1

print("\nBottom 5 movies:")
for name, avg in worst_5:
    print("-", name, "-", avg)

# ======================================================================
# Step 14 - Rating category count (dictionary + loop)
# ======================================================================
category_count = {}
for cat in ["Excellent", "Good", "Average", "Poor"]:
    category_count[cat] = int((df["Rating_Category"] == cat).sum())

print("\nReviews in each category:")
for cat, count in category_count.items():
    print(cat, "->", count)

# ======================================================================
# Step 15 - All charts on ONE page
# ======================================================================
fig, axes = plt.subplots(2, 2, figsize=(14, 9))
fig.suptitle("Movie Review Dashboard", fontsize=16, fontweight="bold")

# Chart 1 (top-left): average rating by genre
axes[0, 0].bar(genre_avg.index, genre_avg.values, color="teal")
axes[0, 0].set_title("Average Rating by Genre")
axes[0, 0].set_ylabel("Average Rating")
axes[0, 0].tick_params(axis="x", rotation=45)

# Chart 2 (top-right): average rating by language
axes[0, 1].bar(language_avg.index, language_avg.values, color="orange")
axes[0, 1].set_title("Average Rating by Language")
axes[0, 1].set_ylabel("Average Rating")

# Chart 3 (bottom-left): rating trend by year
axes[1, 0].plot(year_avg.index, year_avg.values, marker="o", color="crimson")
axes[1, 0].set_title("Average Rating by Year")
axes[1, 0].set_xlabel("Year")
axes[1, 0].set_ylabel("Average Rating")
axes[1, 0].grid(True, linestyle="--", alpha=0.6)

# Chart 4 (bottom-right): how many reviews in each category
axes[1, 1].bar(list(category_count.keys()), list(category_count.values()), color="steelblue")
axes[1, 1].set_title("Reviews by Rating Category")
axes[1, 1].set_ylabel("Number of reviews")

plt.tight_layout(rect=[0, 0, 1, 0.96])
plt.savefig("movie_charts.png")
plt.show()

# ======================================================================
# Step 16 - Save results to Excel spreadsheets
# ======================================================================
df.to_excel("movie_reviews_cleaned.xlsx", index=False)
genre_avg.to_frame("Avg_Rating").to_excel("genre_summary.xlsx")
language_avg.to_frame("Avg_Rating").to_excel("language_summary.xlsx")
movie_avg.to_excel("movie_summary.xlsx")

print("\nExcel files saved: movie_reviews_cleaned.xlsx, genre_summary.xlsx,")
print("language_summary.xlsx, movie_summary.xlsx")

# ======================================================================
# Step 17 - Business observations (answers printed by the program)
# ======================================================================
print("\n========== OBSERVATIONS ==========")
print("1. Highest-rated genre   :", genre_avg.idxmax(), "-", genre_avg.max())
print("2. Lowest-rated genre    :", genre_avg.idxmin(), "-", genre_avg.min())
print("3. Highest-rated language:", language_avg.idxmax(), "-", language_avg.max())
print("4. Best movie            :", best_5[0][0], "-", best_5[0][1])
print("5. Best year             :", year_avg.idxmax(), "-", year_avg.max())
print("6. Rows in raw file / after cleaning:", raw_rows, "/", len(df))
print("   Duplicates removed:", duplicates)
print("   Invalid ratings found:", bad_rating.sum())
print("Note: this is practice data, so these results do not describe real movies.")
