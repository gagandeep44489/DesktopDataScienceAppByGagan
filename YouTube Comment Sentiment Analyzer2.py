"""
Sentiment Analysis Module
Performs sentiment analysis on YouTube comments
"""

from textblob import TextBlob
from enum import Enum

class Sentiment(Enum):
    POSITIVE = "Positive"
    NEUTRAL = "Neutral"
    NEGATIVE = "Negative"

class SentimentAnalyzer:
    """Analyzes sentiment of text using TextBlob"""
    
    def __init__(self):
        self.threshold_positive = 0.1
        self.threshold_negative = -0.1
    
    def analyze(self, text):
        """Analyze sentiment of given text"""
        blob = TextBlob(text)
        polarity = blob.sentiment.polarity
        subjectivity = blob.sentiment.subjectivity
        
        if polarity > self.threshold_positive:
            sentiment = Sentiment.POSITIVE
        elif polarity < self.threshold_negative:
            sentiment = Sentiment.NEGATIVE
        else:
            sentiment = Sentiment.NEUTRAL
        
        return {
            'sentiment': sentiment.value,
            'polarity': polarity,
            'subjectivity': subjectivity,
            'confidence': abs(polarity)
        }
    
    def batch_analyze(self, comments):
        """Analyze multiple comments"""
        results = []
        for comment in comments:
            analysis = self.analyze(comment['text'])
            results.append({
                **comment,
                **analysis
            })
        return results