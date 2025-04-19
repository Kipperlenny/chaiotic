class FuzzyMatcher:
    def __init__(self, choices):
        self.choices = choices

    def get_best_match(self, query):
        best_match = None
        highest_score = 0
        for choice in self.choices:
            score = self._calculate_similarity(query, choice)
            if score > highest_score:
                highest_score = score
                best_match = choice
        return best_match

    def _calculate_similarity(self, query, choice):
        return len(set(query) & set(choice)) / len(set(query) | set(choice))

def generate_full_text_from_corrections(corrections):
    full_text = ""
    for correction in corrections:
        full_text += correction['text']
    return full_text