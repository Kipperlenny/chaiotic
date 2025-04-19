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
    """Generate full text from corrections list.
    
    Args:
        corrections: List of correction dictionaries with 'original' and 'corrected' fields
        
    Returns:
        String containing the corrected full text
    """
    if not corrections or not isinstance(corrections, list):
        return ""
    
    # First try to find corrected_full_text in case it's passed as a dictionary
    if isinstance(corrections, dict) and 'corrected_full_text' in corrections:
        return corrections['corrected_full_text']
    
    # Initialize result text
    full_text = ""
    
    # Ensure corrections are ordered by their occurrence in the original text
    # This assumes the 'original' field contains the exact substring from the original text
    sorted_corrections = sorted(corrections, key=lambda c: c.get('original_index', 0) 
                              if 'original_index' in c else 0)
    
    # Combine the corrected texts with proper spacing
    for i, correction in enumerate(sorted_corrections):
        if not isinstance(correction, dict):
            continue
            
        # Get the corrected text
        if 'corrected' in correction:
            corrected_text = correction['corrected']
        elif 'text' in correction:
            corrected_text = correction['text']
        else:
            # Skip corrections without required fields
            continue
        
        # Add space between sentences if needed
        if i > 0 and corrected_text and full_text and not full_text.endswith((' ', '\n', '\t', '.', '!', '?')) and not corrected_text.startswith((' ', '\n', '\t')):
            full_text += " "
            
        full_text += corrected_text
    
    return full_text