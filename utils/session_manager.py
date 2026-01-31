"""
Session management utilities.

Handles saving and loading optimization sessions to/from JSON files.
"""

import json
import os
from pathlib import Path
from typing import Dict, Any, Optional, List
import logging
from datetime import datetime
import uuid

from config.settings import AppConfig

logger = logging.getLogger(__name__)


class SessionManager:
    """Manager for saving and loading optimization sessions."""
    
    def __init__(self, config: Optional[AppConfig] = None):
        """
        Initialize session manager.
        
        Args:
            config: Application configuration
        """
        self.config = config or AppConfig()
        self.sessions_dir = Path(self.config.SESSIONS_DIR)
        self._ensure_sessions_dir()
    
    def _ensure_sessions_dir(self):
        """Ensure sessions directory exists."""
        self.sessions_dir.mkdir(parents=True, exist_ok=True)
        logger.info(f"Sessions directory: {self.sessions_dir}")
    
    def generate_session_id(self) -> str:
        """
        Generate a unique session ID.
        
        Returns:
            UUID-based session ID
        """
        return str(uuid.uuid4())
    
    def save_session(
        self,
        session_id: str,
        data: Dict[str, Any],
        metadata: Optional[Dict[str, Any]] = None
    ) -> bool:
        """
        Save session data to JSON file.
        
        Args:
            session_id: Unique session identifier
            data: Data to save
            metadata: Optional metadata to include
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Prepare session data
            session_data = {
                'session_id': session_id,
                'timestamp': datetime.now().isoformat(),
                'metadata': metadata or {},
                'data': self._serialize_data(data)
            }
            
            # Write to file
            filepath = self.sessions_dir / f"{session_id}.json"
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(session_data, f, indent=2, ensure_ascii=False)
            
            logger.info(f"Session saved: {session_id}")
            return True
            
        except Exception as e:
            logger.error(f"Error saving session {session_id}: {e}", exc_info=True)
            return False
    
    def load_session(self, session_id: str) -> Optional[Dict[str, Any]]:
        """
        Load session data from JSON file.
        
        Args:
            session_id: Session identifier
            
        Returns:
            Session data or None if not found
        """
        try:
            filepath = self.sessions_dir / f"{session_id}.json"
            
            if not filepath.exists():
                logger.warning(f"Session not found: {session_id}")
                return None
            
            with open(filepath, 'r', encoding='utf-8') as f:
                session_data = json.load(f)
            
            logger.info(f"Session loaded: {session_id}")
            return session_data
            
        except Exception as e:
            logger.error(f"Error loading session {session_id}: {e}", exc_info=True)
            return None
    
    def delete_session(self, session_id: str) -> bool:
        """
        Delete a session file.
        
        Args:
            session_id: Session identifier
            
        Returns:
            True if successful, False otherwise
        """
        try:
            filepath = self.sessions_dir / f"{session_id}.json"
            
            if filepath.exists():
                filepath.unlink()
                logger.info(f"Session deleted: {session_id}")
                return True
            else:
                logger.warning(f"Session not found for deletion: {session_id}")
                return False
                
        except Exception as e:
            logger.error(f"Error deleting session {session_id}: {e}", exc_info=True)
            return False
    
    def list_sessions(self) -> List[Dict[str, Any]]:
        """
        List all saved sessions.
        
        Returns:
            List of session metadata
        """
        try:
            sessions = []
            
            for filepath in self.sessions_dir.glob("*.json"):
                try:
                    with open(filepath, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    
                    sessions.append({
                        'session_id': data.get('session_id'),
                        'timestamp': data.get('timestamp'),
                        'metadata': data.get('metadata', {})
                    })
                except:
                    continue
            
            # Sort by timestamp, newest first
            sessions.sort(
                key=lambda x: x.get('timestamp', ''),
                reverse=True
            )
            
            logger.info(f"Found {len(sessions)} sessions")
            return sessions
            
        except Exception as e:
            logger.error(f"Error listing sessions: {e}", exc_info=True)
            return []
    
    def _serialize_data(self, data: Any) -> Any:
        """
        Recursively serialize data for JSON storage.
        
        Args:
            data: Data to serialize
            
        Returns:
            JSON-serializable data
        """
        import pandas as pd
        import numpy as np
        
        if isinstance(data, dict):
            return {k: self._serialize_data(v) for k, v in data.items()}
        elif isinstance(data, list):
            return [self._serialize_data(item) for item in data]
        elif isinstance(data, pd.DataFrame):
            return {
                '_type': 'dataframe',
                'data': data.to_dict('records'),
                'columns': list(data.columns)
            }
        elif isinstance(data, (np.integer, np.floating)):
            return float(data)
        elif isinstance(data, np.ndarray):
            return data.tolist()
        elif isinstance(data, (datetime,)):
            return data.isoformat()
        else:
            return data
    
    def clean_old_sessions(self, days: int = 30) -> int:
        """
        Delete sessions older than specified days.
        
        Args:
            days: Number of days to keep
            
        Returns:
            Number of sessions deleted
        """
        try:
            cutoff = datetime.now().timestamp() - (days * 24 * 60 * 60)
            deleted = 0
            
            for filepath in self.sessions_dir.glob("*.json"):
                try:
                    stat = filepath.stat()
                    if stat.st_mtime < cutoff:
                        filepath.unlink()
                        deleted += 1
                except:
                    continue
            
            logger.info(f"Deleted {deleted} old sessions")
            return deleted
            
        except Exception as e:
            logger.error(f"Error cleaning old sessions: {e}", exc_info=True)
            return 0
