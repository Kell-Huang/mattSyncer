"""
SKU column detection module.
Automatically identifies SKU columns in source and target files.
"""

import re
from typing import List, Optional, Tuple


class SkuDetector:
    """
    Detects SKU columns in a list of column names.

    Uses pattern matching with confidence scoring to identify
    the most likely SKU column from common naming conventions.
    """

    # Exact match patterns (case insensitive)
    EXACT_MATCHES = [
        "sku",
        "product_id",
        "productid",
        "product-id",
        "seller_id",
        "sellerid",
        "seller-id",
        "ean",
        "ean_code",
        "eancode",
        "ean-code",
        "upc",
        "upc_code",
        "upccode",
        "upc-code",
        "item_id",
        "itemid",
        "item-id",
        "article_id",
        "articleid",
        "article-id",
        "articlenumber",
        "article_number",
        "article-number",
        "part_number",
        "partnumber",
        "part-number",
        "model_number",
        "modelnumber",
        "model-number",
        "style_number",
        "stylenumber",
        "style-number",
        "material_number",
        "materialnumber",
        "material-number",
    ]

    # Fuzzy match patterns with confidence weights
    FUZZY_PATTERNS = [
        (r"\bsku\b", 0.9),  # Exact word 'sku'
        (r"\bid\b", 0.85),  # Exact word 'id'
        (r"sku", 0.7),  # Contains 'sku'
        (r"product.*\bid\b", 0.8),  # 'product' followed by 'id'
        (r"\bid\b.*product", 0.8),  # 'id' followed by 'product'
        (r"ean", 0.7),  # Contains 'ean'
        (r"upc", 0.7),  # Contains 'upc'
        (r"article.*\b(id|number|num)\b", 0.8),  # Article with id/number
        (r"item.*\b(id|number|num|code)\b", 0.7),  # Item with identifier
        (r"part.*\b(id|number|num)\b", 0.6),  # Part with identifier
        (r"model.*\b(id|number|num)\b", 0.6),  # Model with identifier
    ]

    def detect_sku_column(self, columns: List[str]) -> List[str]:
        """Detect SKU columns from a list of column names.

        Args:
            columns: List of column names to search.

        Returns:
            List of potential SKU column names, ordered by confidence
            (highest confidence first).
        """
        if not columns:
            return []

        candidates = []

        for col in columns:
            col_lower = col.lower().strip()
            score = self._calculate_match_score(col_lower)

            if score > 0:
                candidates.append((col, score))

        # Sort by score descending, then by original column order
        candidates.sort(key=lambda x: (-x[1], columns.index(x[0])))

        return [col for col, score in candidates]

    def _calculate_match_score(self, column_name: str) -> float:
        """Calculate a confidence score for a column name being a SKU column.

        Args:
            column_name: Lowercase column name to evaluate.

        Returns:
            Confidence score between 0 and 1.
        """
        # Normalize separators for comparison
        normalized = column_name.replace("-", "_").replace(" ", "_")

        # Check exact matches first (highest confidence)
        if normalized in [m.replace("-", "_") for m in self.EXACT_MATCHES]:
            return 1.0

        # Exact match with original patterns
        if column_name in self.EXACT_MATCHES:
            return 1.0

        # Check for exact match after removing common prefixes
        for prefix in ["source_", "src_", "target_", "tgt_", "file_", "data_"]:
            if column_name.startswith(prefix):
                stripped = column_name[len(prefix) :]
                if stripped in self.EXACT_MATCHES:
                    return 0.9

        # Fuzzy pattern matching
        best_score = 0.0
        for pattern, weight in self.FUZZY_PATTERNS:
            if re.search(pattern, column_name):
                score = weight

                # Bonus for shorter column names (more likely to be a key column)
                if len(column_name) <= 15:
                    score += 0.05

                # Penalty for very long column names
                if len(column_name) > 30:
                    score -= 0.1

                # Penalty for columns that look like descriptions
                if any(
                    word in column_name
                    for word in ["desc", "name", "title", "note"]
                ):
                    score -= 0.2

                best_score = max(best_score, score)

        return min(best_score, 1.0)

    def get_detection_confidence(
        self, columns: List[str]
    ) -> Tuple[Optional[str], str]:
        """Get the detection confidence level for SKU column detection.

        Args:
            columns: List of column names.

        Returns:
            Tuple of (best_match_column, confidence_level) where
            confidence_level is one of: 'high', 'medium', 'low', 'none'.
        """
        candidates = self.detect_sku_column(columns)

        if not candidates:
            return None, "none"

        # Calculate scores for all candidates
        scores = [self._calculate_match_score(c.lower()) for c in candidates]

        if not scores:
            return None, "none"

        # Determine confidence level based on score distribution
        if scores[0] >= 0.9:
            if len(scores) == 1 or scores[0] - scores[1] > 0.3:
                return candidates[0], "high"
            else:
                return candidates[0], "medium"
        elif scores[0] >= 0.7:
            return candidates[0], "medium"
        elif scores[0] >= 0.5:
            return candidates[0], "low"
        else:
            return candidates[0], "none"
