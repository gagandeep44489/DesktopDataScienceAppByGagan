"""
YouTube API Integration
Handles authentication and comment retrieval
"""

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import os

class YouTubeAPIClient:
    SCOPES = ['https://www.googleapis.com/auth/youtube.readonly']
    
    def __init__(self, credentials_file):
        self.credentials_file = credentials_file
        self.service = None
        self.authenticate()
    
    def authenticate(self):
        """Authenticate with YouTube API"""
        creds = None
        
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', self.SCOPES)
        
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.credentials_file, self.SCOPES)
                creds = flow.run_local_server(port=0)
        
        self.service = build('youtube', 'v3', credentials=creds)
    
    def get_video_comments(self, video_id, max_results=100):
        """Retrieve comments from a YouTube video"""
        comments = []
        request = self.service.commentThreads().list(
            part='snippet',
            videoId=video_id,
            maxResults=min(100, max_results),
            textFormat='plainText'
        )
        
        while request and len(comments) < max_results:
            response = request.execute()
            
            for item in response['items']:
                comment = item['snippet']['topLevelComment']['snippet']
                comments.append({
                    'author': comment['authorDisplayName'],
                    'text': comment['textDisplay'],
                    'likes': comment['likeCount'],
                    'timestamp': comment['publishedAt']
                })
            
            if 'nextPageToken' in response:
                request = self.service.commentThreads().list(
                    part='snippet',
                    videoId=video_id,
                    pageToken=response['nextPageToken'],
                    maxResults=min(100, max_results - len(comments)),
                    textFormat='plainText'
                )
            else:
                break
        
        return comments