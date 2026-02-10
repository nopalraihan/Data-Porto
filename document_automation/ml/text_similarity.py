"""
TextSimilarity - ML-powered fuzzy text matching using TF-IDF + Cosine Similarity.

Instead of simple string equality, this module computes semantic similarity
between text fields (names, addresses) to handle:
  - Typos and abbreviations (Jl. vs Jalan, Gg. vs Gang)
  - Missing/extra spaces and punctuation
  - Partial addresses
  - Name variations (with/without titles)
"""

import re
import numpy as np
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity


# Common Indonesian address abbreviations for normalization
ADDRESS_ABBREVIATIONS = {
    r"\bjl\.?\b": "jalan",
    r"\bgg\.?\b": "gang",
    r"\brt\.?\b": "rt",
    r"\brw\.?\b": "rw",
    r"\bkel\.?\b": "kelurahan",
    r"\bkec\.?\b": "kecamatan",
    r"\bkab\.?\b": "kabupaten",
    r"\bno\.?\b": "nomor",
    r"\bblk\.?\b": "blok",
    r"\bjkt\.?\b": "jakarta",
    r"\btmr\.?\b": "timur",
    r"\bbrt\.?\b": "barat",
    r"\bslt\.?\b": "selatan",
    r"\butr\.?\b": "utara",
    r"\bds\.?\b": "desa",
    r"\bperum\.?\b": "perumahan",
    r"\bkomp\.?\b": "kompleks",
}


class TextSimilarity:
    """
    Compute similarity scores between text pairs using TF-IDF character n-grams.

    Methods:
        score(text_a, text_b) -> float  (0.0 to 1.0)
        score_batch(pairs) -> list[float]
        find_best_match(query, candidates) -> (index, score)
    """

    def __init__(self, ngram_range: tuple = (2, 4), analyzer: str = "char_wb"):
        """
        Args:
            ngram_range: Character n-gram range for TF-IDF.
            analyzer: 'char_wb' for character n-grams within word boundaries.
        """
        self.ngram_range = ngram_range
        self.analyzer = analyzer

    @staticmethod
    def _normalize_text(text: str) -> str:
        """Normalize text for comparison."""
        if not text:
            return ""
        text = str(text).lower().strip()
        # Expand abbreviations
        for pattern, replacement in ADDRESS_ABBREVIATIONS.items():
            text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)
        # Remove excess whitespace and punctuation noise
        text = re.sub(r"[,\.\-\/\\]+", " ", text)
        text = re.sub(r"\s+", " ", text)
        return text.strip()

    def score(self, text_a: str, text_b: str) -> float:
        """
        Compute similarity score between two texts.

        Returns:
            float between 0.0 (no similarity) and 1.0 (identical).
        """
        a = self._normalize_text(text_a)
        b = self._normalize_text(text_b)

        if not a or not b:
            return 0.0
        if a == b:
            return 1.0

        try:
            vectorizer = TfidfVectorizer(
                analyzer=self.analyzer,
                ngram_range=self.ngram_range,
            )
            tfidf_matrix = vectorizer.fit_transform([a, b])
            sim = cosine_similarity(tfidf_matrix[0:1], tfidf_matrix[1:2])[0][0]
            return round(float(sim), 4)
        except ValueError:
            return 0.0

    def score_batch(self, pairs: list[tuple[str, str]]) -> list[float]:
        """
        Compute similarity for multiple text pairs efficiently.

        Args:
            pairs: list of (text_a, text_b) tuples

        Returns:
            list of similarity scores
        """
        if not pairs:
            return []

        all_texts = []
        for a, b in pairs:
            all_texts.extend([self._normalize_text(a), self._normalize_text(b)])

        try:
            vectorizer = TfidfVectorizer(
                analyzer=self.analyzer,
                ngram_range=self.ngram_range,
            )
            tfidf_matrix = vectorizer.fit_transform(all_texts)

            scores = []
            for i in range(0, len(all_texts), 2):
                sim = cosine_similarity(tfidf_matrix[i:i+1], tfidf_matrix[i+1:i+2])[0][0]
                scores.append(round(float(sim), 4))
            return scores
        except ValueError:
            return [0.0] * len(pairs)

    def find_best_match(self, query: str, candidates: list[str]) -> tuple[int, float]:
        """
        Find the most similar candidate to the query.

        Returns:
            (best_index, best_score)
        """
        if not candidates:
            return (-1, 0.0)

        query_norm = self._normalize_text(query)
        cand_norms = [self._normalize_text(c) for c in candidates]

        all_texts = [query_norm] + cand_norms

        try:
            vectorizer = TfidfVectorizer(
                analyzer=self.analyzer,
                ngram_range=self.ngram_range,
            )
            tfidf_matrix = vectorizer.fit_transform(all_texts)
            sims = cosine_similarity(tfidf_matrix[0:1], tfidf_matrix[1:])[0]
            best_idx = int(np.argmax(sims))
            return (best_idx, round(float(sims[best_idx]), 4))
        except ValueError:
            return (0, 0.0)

    def classify_match(self, score: float) -> str:
        """Classify a similarity score into a human-readable category."""
        if score >= 0.95:
            return "EXACT"
        elif score >= 0.80:
            return "HIGH"
        elif score >= 0.60:
            return "MEDIUM"
        elif score >= 0.40:
            return "LOW"
        else:
            return "NO_MATCH"
