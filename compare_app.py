from flask import Flask, render_template, request
import pandas as pd
import os
import re

app = Flask(__name__)

GOLD_FILE = r"C:\mythos\Gold-Standard\mythos_gold_standard.xlsx"


def clean_text(text):
    if pd.isna(text):
        return ""
    text = str(text).lower()
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_domain(text):
    return clean_text(text)


def compare_answers(client_answer, gold_answer):
    client = clean_text(client_answer)
    gold = clean_text(gold_answer)

    if not client:
        return "High", 0

    stop_words = {
        "with", "that", "this", "from", "they", "have", "into",
        "such", "using", "all", "and", "the", "for", "are", "can",
        "you", "your", "what", "which", "how", "does", "been"
    }

    gold_words = set(gold.split())
    client_words = set(client.split())

    important_words = {
        word for word in gold_words
        if len(word) > 3 and word not in stop_words
    }

    if not important_words:
        return "High", 0

    matched_words = important_words.intersection(client_words)
    score = len(matched_words) / len(important_words)

    if score >= 0.65:
        return "Low", round(score * 100, 2)
    elif score >= 0.30:
        return "Medium", round(score * 100, 2)
    else:
        return "High", round(score * 100, 2)


def load_client_file(uploaded_file):
    df = pd.read_excel(uploaded_file, header=None)

    # Your file format:
    # 0 = Timestamp
    # 1 = Domain
    # 2 = Question
    # 3 = Answer
    # 4 = Risk
    df = df.iloc[:, :5]
    df.columns = ["Timestamp", "Domain", "Question", "Answer", "Risk"]

    return df


def load_gold_file():
    gold_df = pd.read_excel(GOLD_FILE)

    gold_df.columns = [str(col).strip() for col in gold_df.columns]

    gold_map = {}

    for _, row in gold_df.iterrows():
        domain = normalize_domain(row["Domain"])
        answer = row["Gold Standard Answer"]
        gold_map[domain] = answer

    return gold_map


@app.route("/", methods=["GET", "POST"])
def index():
    results = []
    summary = None
    error = None

    if request.method == "POST":
        uploaded_file = request.files.get("client_file")

        if not uploaded_file:
            error = "Please select a client Excel file."
            return render_template("compare.html", results=results, summary=summary, error=error)

        if not os.path.exists(GOLD_FILE):
            error = f"Gold standard file not found: {GOLD_FILE}"
            return render_template("compare.html", results=results, summary=summary, error=error)

        try:
            client_df = load_client_file(uploaded_file)
            gold_map = load_gold_file()

            low_count = 0
            medium_count = 0
            high_count = 0

            for _, row in client_df.iterrows():
                domain = row["Domain"]
                question = row["Question"]
                client_answer = row["Answer"]
                selected_risk = row["Risk"]

                gold_answer = gold_map.get(normalize_domain(domain), "")

                risk_level, match_score = compare_answers(client_answer, gold_answer)

                if risk_level == "Low":
                    low_count += 1
                elif risk_level == "Medium":
                    medium_count += 1
                else:
                    high_count += 1

                results.append({
                    "domain": domain,
                    "question": question,
                    "client_answer": client_answer,
                    "gold_answer": gold_answer,
                    "selected_risk": selected_risk,
                    "match_score": match_score,
                    "risk_level": risk_level
                })

            total = low_count + medium_count + high_count

            summary = {
                "total": total,
                "low": low_count,
                "medium": medium_count,
                "high": high_count,
                "high_risk_percent": round((high_count / total) * 100, 2) if total else 0
            }

        except Exception as e:
            error = str(e)

    return render_template("compare.html", results=results, summary=summary, error=error)


if __name__ == "__main__":
    app.run(debug=True, port=5002)