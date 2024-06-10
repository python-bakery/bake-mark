import spacy

SLIDE_DELAY = 1
NARRATION_DELAY = 1
MAX_LINE_LENGTH = 80
MIN_LINE_LENGTH = 40

nlp_sentencizer = spacy.blank("en")
nlp_sentencizer.add_pipe("sentencizer")

def find_nearest_space(text: str, index: int, extent=5) -> int:
    for i in range(0, extent):
        if text[index + i] == " ":
            return index + i
        if text[index - i] == " ":
            return index - i
    return index

def as_time(seconds: int) -> str:
    minutes = seconds // 60
    return f"{minutes:02}:{seconds % 60:02}.000"

def split_sentences(text: str, max_line_length=MAX_LINE_LENGTH, min_line_length = MIN_LINE_LENGTH) -> list[str]:
    tokens = nlp_sentencizer(text)
    sentences = [str(sent).strip() for sent in tokens.sents]
    for sentence in sentences:
        split_sentences = [sentence]
        while len(split_sentences[-1]) > max_line_length:
            current_sentence = split_sentences.pop()
            first, rest = current_sentence[:max_line_length], current_sentence[max_line_length:]
            if len(rest) < min_line_length:
                halfway_point = find_nearest_space(current_sentence, len(current_sentence) // 2)
                first, rest = current_sentence[:halfway_point], current_sentence[halfway_point:]
            split_sentences.extend([first, rest])
        yield from split_sentences
        
        
#print(list(split_sentences("This is a test of the sentence splitter. It should split this sentence into two.", 5, 2)))

def make_captions(transcript: list[str], durations: list[int]) -> list[str]:
    yield "WEBVTT\n"
    current_time = SLIDE_DELAY
    for text, duration in zip(transcript, durations):
        text = text.replace('\n', ' ')    
        sentences = list(split_sentences(text))
        sentence_lengths = [len(sent) for sent in sentences]
        total_length = sum(sentence_lengths)
        sentence_durations = [round(duration * (length / total_length)) for length in sentence_lengths]
        time_offset = current_time
        sentence_durations[-1] += SLIDE_DELAY
        for sentence, sentence_duration in zip(sentences, sentence_durations):
            yield f"{as_time(time_offset)} --> {as_time(time_offset + sentence_duration)}"
            yield f"{sentence.strip()}\n"
            time_offset += sentence_duration
        current_time += duration + SLIDE_DELAY
        
        
#print("\n".join(make_captions(["This is a test of the sentence splitter. It should split this sentence into two.",
#                          "However are you doing today, this will be a very long sentence indeed, twice as long in fact."], [5, 10])))
