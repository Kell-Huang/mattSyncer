"""
Column name matching algorithm.
Implements direct prefix-stripping match and hybrid similarity scoring
for matching source columns to target columns.
"""

import re
from collections import Counter
from dataclasses import dataclass
from typing import Dict, List, Optional, Set, Tuple


@dataclass
class MatchResult:
    """Result of a single column matching operation."""

    source: str
    target: Optional[str]
    match_type: str  # 'direct', 'indirect', or 'none'
    score: Optional[float] = None
    details: Optional[Dict] = None


class ColumnMatcher:
    """
    Matches source columns to target columns using prefix stripping
    and hybrid similarity scoring.

    The algorithm:
    1. Attempts prefix-stripping exact match (e.g., 'category' -> 'home24_category')
    2. Falls back to hybrid similarity scoring for fuzzy matches
    3. Returns match results with confidence information
    """

    # Weights for token-level similarity components
    TOKEN_LCS_WEIGHT = 0.4
    TOKEN_EDIT_WEIGHT = 0.2
    TOKEN_OVERLAP_WEIGHT = 0.4

    # Weights for compound vs single-token column names
    COMPOUND_TOKEN_WEIGHT = 0.8
    COMPOUND_CHAR_WEIGHT = 0.2

    SINGLE_TOKEN_WEIGHT = 0.2
    SINGLE_CHAR_WEIGHT = 0.8

    # Penalty multiplier when units differ
    UNIT_PENALTY = 0.3

    # Minimum score threshold for indirect match
    MIN_INDIRECT_SCORE = 0.3

    def match_columns(
        self, source_columns: List[str], target_columns: List[str]
    ) -> Tuple[List[Dict], Dict[str, str]]:
        """Match all source columns to target columns.

        Args:
            source_columns: List of source column names.
            target_columns: List of target column names.

        Returns:
            Tuple of (match_results_list, prefix_info_dict).
            prefix_info contains 'most_common_prefix', 'prefix_counter',
            and 'all_prefixes'.
        """
        results = []
        prefix_counter = Counter()

        for source_col in source_columns:
            result = self._match_single_column(source_col, target_columns)

            # Extract prefix from direct matches
            prefix = None
            if result.match_type == "direct" and result.target:
                prefix = self._extract_prefix(source_col, result.target)
                if prefix:
                    prefix_counter[prefix] += 1

            results.append(
                {
                    "source": source_col,
                    "target": result.target,
                    "match_type": result.match_type,
                    "score": result.score,
                    "prefix": prefix,
                    "all_targets": target_columns,
                }
            )

        # Determine the most common prefix
        most_common_prefix = None
        if prefix_counter:
            most_common_prefix = prefix_counter.most_common(1)[0][0]

        prefix_info = {
            "most_common_prefix": most_common_prefix,
            "prefix_counter": dict(prefix_counter),
            "all_prefixes": list(prefix_counter.keys()),
        }

        return results, prefix_info

    def _extract_prefix(self, source: str, target: str) -> Optional[str]:
        """Extract the prefix stripped from target column in direct match.

        For example, if source='category' and target='home24_category',
        returns 'home24_'.

        Args:
            source: Source column name.
            target: Target column name.

        Returns:
            The prefix string with trailing underscore, or None.
        """
        source_lower = source.lower().strip()
        target_lower = target.lower().strip()

        # Split target by underscore
        parts = target_lower.split("_")
        if len(parts) > 1:
            # Remove first part and check if remainder matches source
            stripped = "_".join(parts[1:])
            if stripped == source_lower:
                return parts[0] + "_"

        return None

    def generate_new_column_name(
        self, source_column: str, prefix: str, existing_columns: List[str]
    ) -> Optional[str]:
        """Generate a new column name by prepending prefix to source column.

        Args:
            source_column: Source column name.
            prefix: Prefix to prepend (e.g., 'home24_').
            existing_columns: List of existing target column names.

        Returns:
            New column name, or None if it would conflict with existing column.
        """
        if prefix:
            new_name = prefix + source_column
        else:
            new_name = source_column

        # Check for case-insensitive conflicts
        existing_lower = [c.lower() for c in existing_columns]
        if new_name.lower() in existing_lower:
            return None

        return new_name

    def _match_single_column(
        self, source: str, targets: List[str]
    ) -> MatchResult:
        """Match a single source column against all target columns.

        Args:
            source: Source column name.
            targets: List of target column names.

        Returns:
            MatchResult with the best match information.
        """
        # Step 0: Try prefix-stripping direct match
        direct_match = self._direct_match(source, targets)
        if direct_match:
            return MatchResult(
                source=source, target=direct_match, match_type="direct"
            )

        # Step 1: Hybrid similarity scoring
        if not targets:
            return MatchResult(source=source, target=None, match_type="none")

        best_target = None
        best_score = -1.0

        for target in targets:
            score = self._hybrid_score(source, target)
            if score > best_score:
                best_score = score
                best_target = target

        # Only return indirect match if score meets threshold
        if best_score >= self.MIN_INDIRECT_SCORE:
            return MatchResult(
                source=source,
                target=best_target,
                match_type="indirect",
                score=best_score,
            )
        else:
            return MatchResult(
                source=source, target=None, match_type="none", score=best_score
            )

    def _direct_match(self, source: str, targets: List[str]) -> Optional[str]:
        """Attempt prefix-stripping direct match.

        Removes the first underscore-separated segment from target columns
        and checks for exact match with source.

        Args:
            source: Source column name.
            targets: List of target column names.

        Returns:
            First matching target column, or None.
        """
        source_lower = source.lower().strip()

        for target in targets:
            target_lower = target.lower().strip()

            # Split by underscore and remove first element
            parts = target_lower.split("_")
            if len(parts) > 1:
                stripped = "_".join(parts[1:])
                if stripped and stripped == source_lower:
                    return target

        return None

    def _normalize(self, column_name: str) -> Tuple[List[str], Optional[str]]:
        """Normalize a column name into tokens and optional unit.

        Steps:
        1. Convert to lowercase, strip whitespace
        2. Extract unit from trailing parentheses
        3. Split by underscores and hyphens
        4. Split camelCase and digit boundaries

        Args:
            column_name: Raw column name string.

        Returns:
            Tuple of (tokens_list, unit_string_or_None).
        """
        col = column_name.lower().strip()

        # Extract unit from trailing parentheses
        unit = None
        unit_match = re.search(r"[（(](.+?)[）)]$", col)
        if unit_match:
            unit = unit_match.group(1).strip()
            col = col[: unit_match.start()].strip()

        # Split by underscores and hyphens
        segments = re.split(r"[_-]", col)

        # Split each segment by case boundaries and digits
        tokens = []
        for segment in segments:
            sub_tokens = re.findall(r"[a-z]+|[0-9]+", segment.lower())
            tokens.extend(sub_tokens)

        # Filter empty tokens
        tokens = [t for t in tokens if t]

        return tokens, unit

    def _lcs_sim(self, seq_a: List[str], seq_b: List[str]) -> float:
        """Calculate Longest Common Subsequence similarity.

        Formula: sim = 2 * LCS_length / (len(seq_a) + len(seq_b))

        Args:
            seq_a: First sequence of tokens/chars.
            seq_b: Second sequence of tokens/chars.

        Returns:
            Similarity score between 0.0 and 1.0.
        """
        if not seq_a and not seq_b:
            return 1.0
        if not seq_a or not seq_b:
            return 0.0

        m, n = len(seq_a), len(seq_b)

        # Dynamic programming table for LCS
        dp = [[0] * (n + 1) for _ in range(m + 1)]

        for i in range(1, m + 1):
            for j in range(1, n + 1):
                if seq_a[i - 1] == seq_b[j - 1]:
                    dp[i][j] = dp[i - 1][j - 1] + 1
                else:
                    dp[i][j] = max(dp[i - 1][j], dp[i][j - 1])

        lcs_length = dp[m][n]
        return 2 * lcs_length / (len(seq_a) + len(seq_b))

    def _edit_sim(self, seq_a: List[str], seq_b: List[str]) -> float:
        """Calculate Levenshtein edit distance similarity.

        Formula: sim = 1 - edit_distance / max(len(seq_a), len(seq_b))
        Operations: insertion, deletion, substitution (cost=1 each).

        Args:
            seq_a: First sequence of tokens/chars.
            seq_b: Second sequence of tokens/chars.

        Returns:
            Similarity score between 0.0 and 1.0.
        """
        if not seq_a and not seq_b:
            return 1.0
        if not seq_a or not seq_b:
            return 0.0

        m, n = len(seq_a), len(seq_b)

        # Dynamic programming table for edit distance
        dp = [[0] * (n + 1) for _ in range(m + 1)]

        # Initialize base cases
        for i in range(m + 1):
            dp[i][0] = i
        for j in range(n + 1):
            dp[0][j] = j

        # Fill DP table
        for i in range(1, m + 1):
            for j in range(1, n + 1):
                cost = 0 if seq_a[i - 1] == seq_b[j - 1] else 1
                dp[i][j] = min(
                    dp[i - 1][j] + 1,  # Deletion
                    dp[i][j - 1] + 1,  # Insertion
                    dp[i - 1][j - 1] + cost,  # Substitution
                )

        edit_distance = dp[m][n]
        max_len = max(m, n)
        return 1.0 - edit_distance / max_len

    def _overlap_coef(self, set_a: Set[str], set_b: Set[str]) -> float:
        """Calculate overlap coefficient between two sets.

        Formula: overlap = |intersection| / min(|set_a|, |set_b|)

        Args:
            set_a: First set of tokens.
            set_b: Second set of tokens.

        Returns:
            Overlap coefficient between 0.0 and 1.0.
        """
        if not set_a or not set_b:
            return 0.0

        intersection = len(set_a & set_b)
        min_size = min(len(set_a), len(set_b))

        return intersection / min_size if min_size > 0 else 0.0

    def _hybrid_score(self, col_a: str, col_b: str) -> float:
        """Calculate hybrid similarity score between two column names.

        Combines token-level and character-level similarity with
        dynamic weighting based on column name complexity.

        Args:
            col_a: First column name.
            col_b: Second column name.

        Returns:
            Similarity score between 0.0 and 1.0.
        """
        # Step 1: Normalize both column names
        tokens_a, unit_a = self._normalize(col_a)
        tokens_b, unit_b = self._normalize(col_b)

        # Step 2: Token-level similarity (three indicators)
        token_lcs = self._lcs_sim(tokens_a, tokens_b)
        token_edit = self._edit_sim(tokens_a, tokens_b)
        token_overlap = self._overlap_coef(set(tokens_a), set(tokens_b))

        token_score = (
            self.TOKEN_LCS_WEIGHT * token_lcs
            + self.TOKEN_EDIT_WEIGHT * token_edit
            + self.TOKEN_OVERLAP_WEIGHT * token_overlap
        )

        # Step 3: Character-level similarity (two indicators)
        char_a = list(col_a.lower())
        char_b = list(col_b.lower())

        char_lcs = self._lcs_sim(char_a, char_b)
        char_edit = self._edit_sim(char_a, char_b)
        char_score = (char_lcs + char_edit) / 2.0

        # Step 4: Dynamic weight fusion
        if len(tokens_a) > 1 or len(tokens_b) > 1:
            # At least one is compound: token-level dominant
            score = (
                self.COMPOUND_TOKEN_WEIGHT * token_score
                + self.COMPOUND_CHAR_WEIGHT * char_score
            )
        else:
            # Both are single-word: character-level dominant
            score = (
                self.SINGLE_TOKEN_WEIGHT * token_score
                + self.SINGLE_CHAR_WEIGHT * char_score
            )

        # Step 5: Unit penalty
        if unit_a is not None and unit_b is not None and unit_a != unit_b:
            score *= self.UNIT_PENALTY

        return min(score, 1.0)
