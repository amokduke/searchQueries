import argparse
import csv
import re
from difflib import SequenceMatcher

import pandas as pd


def normalise_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip().lower()


def normalise_phone(value) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\D", "", str(value))


def normalise_postal(value) -> str:
    if pd.isna(value):
        return ""
    digits = re.sub(r"\D", "", str(value))
    return digits.zfill(6) if digits else ""


def similarity(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()


def score_candidate(query: dict, member: dict) -> tuple[float, list[str]]:
    score = 0.0
    reasons = []

    q_name = normalise_text(query.get("query_name", ""))
    q_email = normalise_text(query.get("query_email", ""))
    q_tel = normalise_phone(query.get("query_telephone", ""))
    q_postal = normalise_postal(query.get("query_postal_code", ""))

    m_name = normalise_text(member.get("name", ""))
    m_email = normalise_text(member.get("email", ""))
    m_tel = normalise_phone(member.get("telephone", ""))
    m_postal = normalise_postal(member.get("postal_code", ""))

    name_ratio = similarity(q_name, m_name)

    has_strong_non_postal_match = False
    has_postal_match = False
    has_name_signal = False

    # Exact email match
    if q_email and m_email and q_email == m_email:
        score += 100
        reasons.append("exact_email")
        has_strong_non_postal_match = True

    # Exact phone match
    if q_tel and m_tel and q_tel == m_tel:
        score += 100
        reasons.append("exact_telephone")
        has_strong_non_postal_match = True

    # Partial telephone safeguard
    if q_tel and m_tel and q_tel in m_tel and q_tel != m_tel:
        score += 20
        reasons.append("partial_telephone")
        has_strong_non_postal_match = True

    # Partial email safeguard
    if q_email and m_email and q_email in m_email and q_email != m_email:
        score += 20
        reasons.append("partial_email")
        has_strong_non_postal_match = True

    # Exact name match
    if q_name and m_name and q_name == m_name:
        score += 80
        reasons.append("exact_name")
        has_name_signal = True
    elif q_name and m_name and name_ratio >= 0.75:
        # Fuzzy name only if there is another useful match already,
        # or if postal code is present and the name is reasonably close
        if has_strong_non_postal_match or has_postal_match or (q_postal and m_postal and q_postal == m_postal):
            fuzzy_points = round(name_ratio * 60, 2)
            score += fuzzy_points
            reasons.append(f"fuzzy_name:{name_ratio:.2f}")
            has_name_signal = True

    # Postal code match
    if q_postal and m_postal and q_postal == m_postal:
        has_postal_match = True

        # Postal code only adds score if there is some other evidence
        if has_strong_non_postal_match or has_name_signal:
            score += 30
            reasons.append("exact_postal_code")

    return score, reasons


def load_members(member_file: str) -> pd.DataFrame:
    df = pd.read_csv(member_file, dtype=str).fillna("")
    required_cols = [
        "_src_idx", "member_id", "name", "year of birth", "email", "telephone",
        "membership", "ethnicity", "postal_code", "unit address", "status",
        "diagnosis of first dependent", "diagnosis of second dependent",
        "created on", "modified on", "postal6", "lat", "lon"
    ]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Missing columns in member file: {missing}")
    return df


def load_queries(query_file: str) -> pd.DataFrame:
    df = pd.read_csv(query_file, dtype=str).fillna("")
    required_cols = ["query_name", "query_email", "query_telephone", "query_postal_code"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Missing columns in query file: {missing}")
    return df


def find_matches(members_df: pd.DataFrame, queries_df: pd.DataFrame, top_n: int, min_score: float) -> pd.DataFrame:
    output_rows = []

    member_records = members_df.to_dict(orient="records")

    for query_idx, query_row in queries_df.iterrows():
        query = query_row.to_dict()
        scored = []

        for member in member_records:
            score, reasons = score_candidate(query, member)
            if score >= min_score:
                scored.append((score, reasons, member))

        scored.sort(key=lambda x: x[0], reverse=True)

        if not scored:
            output_rows.append({
                "query_index": query_idx + 1,
                "query_name": query.get("query_name", ""),
                "query_email": query.get("query_email", ""),
                "query_telephone": query.get("query_telephone", ""),
                "query_postal_code": query.get("query_postal_code", ""),
                "match_rank": "",
                "match_score": "",
                "match_reasons": "no_match",
                "member_id": "",
                "name": "",
                "email": "",
                "telephone": "",
                "postal_code": "",
                "membership": "",
                "status": "",
            })
            continue

        for rank, (score, reasons, member) in enumerate(scored[:top_n], start=1):
            output_rows.append({
                "query_index": query_idx + 1,
                "query_name": query.get("query_name", ""),
                "query_email": query.get("query_email", ""),
                "query_telephone": query.get("query_telephone", ""),
                "query_postal_code": query.get("query_postal_code", ""),
                "match_rank": rank,
                "match_score": round(score, 2),
                "match_reasons": ";".join(reasons),
                "member_id": member.get("member_id", ""),
                "name": member.get("name", ""),
                "email": member.get("email", ""),
                "telephone": member.get("telephone", ""),
                "postal_code": member.get("postal_code", ""),
                "membership": member.get("membership", ""),
                "status": member.get("status", ""),
            })

    return pd.DataFrame(output_rows)


def main():
    parser = argparse.ArgumentParser(
        description="Search member records and surface likely member_id matches."
    )
    parser.add_argument(
        "--members",
        default="members_with_constituencyDisplay_with_mp.csv",
        help="Path to the members CSV file"
    )
    parser.add_argument(
        "--queries",
        default="search_queries.csv",
        help="Path to the query CSV file with columns: query_name, query_email, query_telephone, query_postal_code"
    )
    parser.add_argument(
        "--output",
        default="member_search_results.csv",
        help="Path to the output CSV file"
    )
    parser.add_argument(
        "--top-n",
        type=int,
        default=5,
        help="Number of top possible matches to output for each query"
    )
    parser.add_argument(
        "--min-score",
        type=float,
        default=30.0,
        help="Minimum score threshold to include a candidate"
    )

    args = parser.parse_args()

    members_df = load_members(args.members)
    queries_df = load_queries(args.queries)
    results_df = find_matches(members_df, queries_df, args.top_n, args.min_score)

    results_df.to_csv(args.output, index=False, quoting=csv.QUOTE_MINIMAL)
    print(f"Done. Results written to: {args.output}")


if __name__ == "__main__":
    main()