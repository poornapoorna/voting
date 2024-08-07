import openpyxl
from collections import defaultdict

# Task 1: Implement the generate_preferences function

def generate_preferences(values):
    preferences = defaultdict(list)

    for agent, row in enumerate(values, start=1):
        preference_order = sorted(range(len(row)), key=lambda x: (row[x], -x))
        preferences[agent] = preference_order

    return preferences

# Task 2: Implement the functions

def dictatorship(preferences, agent):
    try:
        return preferences[agent][0]
    except KeyError:
        raise ValueError("Invalid agent")

def scoring_rule(preferences, score_vector, tie_break):
    try:
        # Your implementation here
        pass
    except IndexError:
        print("Incorrect input")
        return False

def plurality(preferences, tie_break):
    scores = defaultdict(int)

    for agent_preference in preferences.values():
        top_choice = agent_preference[0]
        scores[top_choice] += 1

    # Find the alternative with the maximum score
    max_score = max(scores.values())
    possible_winners = [alternative for alternative, score in scores.items() if score == max_score]

    # Apply tie-breaking rules
    if tie_break == "max":
        return max(possible_winners)
    elif tie_break == "min":
        return min(possible_winners)
    elif isinstance(tie_break, int):
        try:
            return max(possible_winners, key=lambda x: preferences[tie_break].index(x))
        except ValueError:
            raise ValueError("Invalid agent for tie-breaking")



def veto(preferences, tie_break):
    """
    Implements the Veto voting rule with error handling, tie-breaking, and preference validation.

    Args:
        preferences (dict): A dictionary representing the preference profile, where keys are agents
                            (1, 2, ...) and values are lists of preferred alternatives in descending order.
        tie_break (str or int): Tie-breaking rule ("max", "min", or agent index).

    Returns:
        int: The winning alternative according to the Veto rule.

    Raises:
        ValueError: If preferences are invalid or if the tie-breaking option is not supported.
    """

    # Validate input
    if not isinstance(preferences, dict):
        raise ValueError("Preferences must be a dictionary.")

    for agent, pref_list in preferences.items():
        if not isinstance(agent, int) or not isinstance(pref_list, list):
            raise ValueError("Invalid preference format.")

    # Initialize scores
    scores = {alt: len(preferences) for alt in set().union(*preferences.values())}  # All scores start with maximum

    # Apply vetoes
    for agent, pref_list in preferences.items():
        scores[pref_list[-1]] -= 1  # Subtract 1 point from the last-ranked alternative

    # Identify potential winners
    winners = [alt for alt, score in scores.items() if score == max(scores.values())]

    # Handle ties
    if len(winners) > 1:
        if tie_break == "max":
            return max(winners)  # Select the alternative with the highest numerical value
        elif tie_break == "min":
            return min(winners)  # Select the alternative with the lowest numerical value
        elif isinstance(tie_break, int) and 1 <= tie_break <= len(preferences):
            return preferences[tie_break][0]  # Select the top preference of the specified agent
        else:
            raise ValueError("Invalid tie-breaking option.")
    else:
        return winners[0]

def borda(preferences, tie_break):
    scores = defaultdict(int)

    for agent_preference in preferences.values():
        for i, alternative in enumerate(agent_preference):
            scores[alternative] += len(agent_preference) - i - 1  # Assign m-j points to the j-th choice

    # Find the alternative with the maximum score
    max_score = max(scores.values())
    possible_winners = [alternative for alternative, score in scores.items() if score == max_score]

    # Apply tie-breaking rules
    if tie_break == "max":
        return max(possible_winners)
    elif tie_break == "min":
        return min(possible_winners)
    elif isinstance(tie_break, int):
        try:
            return max(possible_winners, key=lambda x: preferences[tie_break].index(x))
        except ValueError:
            raise ValueError("Invalid agent for tie-breaking")

def harmonic(preferences, tie_break):
    scores = defaultdict(float)

    for agent_preference in preferences.values():
        for i, alternative in enumerate(agent_preference):
            scores[alternative] += 1 / (i + 1)  # Assign 1/j points to the j-th choice

    # Find the alternative with the maximum score
    max_score = max(scores.values())
    possible_winners = [alternative for alternative, score in scores.items() if score == max_score]

    # Apply tie-breaking rules
    if tie_break == "max":
        return max(possible_winners)
    elif tie_break == "min":
        return min(possible_winners)
    elif isinstance(tie_break, int):
        try:
            return max(possible_winners, key=lambda x: preferences[tie_break].index(x))
        except ValueError:
            raise ValueError("Invalid agent for tie-breaking")

def STV(preferences, tie_break):
    alternatives = set(alternative for preference_order in preferences.values() for alternative in preference_order)
    eliminated = set()

    while len(alternatives) > 1:
        scores = defaultdict(int)

        for agent_preference in preferences.values():
            for i, alternative in enumerate(agent_preference):
                if alternative not in eliminated:
                    scores[alternative] += 1
                    break  # Only count the first choice that is still in the game

        min_score = min(scores.values())
        eliminated.update(alternative for alternative, score in scores.items() if score == min_score)

    # Find the alternative with the maximum score
    winner = max(alternatives)

    # Apply tie-breaking rules
    if tie_break == "max":
        return winner
    elif tie_break == "min":
        return min(alternatives)
    elif isinstance(tie_break, int):
        try:
            return max(alternatives, key=lambda x: preferences[tie_break].index(x))
        except ValueError:
            raise ValueError("Invalid agent for tie-breaking")

def range_voting(values, tie_break):
    # Calculate the sum of valuations for each alternative
    sums = [sum(row) for row in values]

    # Find the alternative with the maximum sum
    winner = sums.index(max(sums)) + 1  # Add 1 to convert from 0-based to 1-based indexing

    # Apply tie-breaking rules
    if tie_break == "max":
        return winner
    elif tie_break == "min":
        return sums.index(min(sums)) + 1
    elif isinstance(tie_break, int):
        try:
            return max(range(1, len(sums) + 1), key=lambda x: values[tie_break - 1][x - 1])
        except IndexError:
            raise ValueError("Invalid agent for tie-breaking")
