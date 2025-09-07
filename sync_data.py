#!/usr/bin/env python3
"""
סקריפט סינכרון נתונים אוטומטי עם Git
"""

import subprocess
import os
import sys
import json
from datetime import datetime

class GitSyncManager:
    def __init__(self, config_file="config.json"):
        self.repo_path = os.getcwd()
        self.config_file = config_file
        self.config = self.load_config()
        
    def load_config(self):
        """טעינת הגדרות"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return {}
        except Exception:
            return {}
    
    def save_config(self):
        """שמירת הגדרות"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
            return True
        except Exception:
            return False
        
    def run_git_command(self, command):
        """הרצת פקודת Git"""
        try:
            result = subprocess.run(
                command, 
                cwd=self.repo_path, 
                capture_output=True, 
                text=True, 
                check=True
            )
            return True, result.stdout
        except subprocess.CalledProcessError as e:
            return False, e.stderr
    
    def is_auto_sync_enabled(self):
        """בדיקה אם סינכרון אוטומטי מופעל"""
        return self.config.get("git", {}).get("auto_sync_enabled", False)
    
    def sync_data(self, force=False):
        """סינכרון מלא של הנתונים"""
        if not force and not self.is_auto_sync_enabled():
            print("ℹ️  סינכרון אוטומטי מושבת")
            return True
            
        print("🔄 מתחיל סינכרון נתונים...")
        
        # 1. משיכת שינויים
        print("📥 מושך שינויים מהמאגר...")
        branch = self.config.get("git", {}).get("branch", "main")
        success, output = self.run_git_command(["git", "pull", "origin", branch])
        if not success:
            print(f"❌ שגיאה במשיכת שינויים: {output}")
            return False
        
        # 2. הוספת שינויים
        print("📝 מוסיף שינויים...")
        success, output = self.run_git_command(["git", "add", "."])
        if not success:
            print(f"❌ שגיאה בהוספת שינויים: {output}")
            return False
        
        # 3. בדיקה אם יש שינויים
        success, output = self.run_git_command(["git", "status", "--porcelain"])
        if not success or not output.strip():
            print("ℹ️  אין שינויים חדשים")
            # עדכון זמן סינכרון אחרון
            self.update_last_sync()
            return True
        
        # 4. יצירת commit
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        commit_message = f"סינכרון נתונים אוטומטי - {timestamp}"
        
        print("💾 יוצר commit...")
        success, output = self.run_git_command(["git", "commit", "-m", commit_message])
        if not success:
            print(f"❌ שגיאה ביצירת commit: {output}")
            return False
        
        # 5. דחיפה למאגר
        print("📤 דוחף שינויים למאגר...")
        success, output = self.run_git_command(["git", "push", "origin", branch])
        if not success:
            print(f"❌ שגיאה בדחיפת שינויים: {output}")
            return False
        
        # עדכון זמן סינכרון אחרון
        self.update_last_sync()
        print("✅ סינכרון הושלם בהצלחה!")
        return True
    
    def update_last_sync(self):
        """עדכון זמן סינכרון אחרון"""
        if "git" not in self.config:
            self.config["git"] = {}
        self.config["git"]["last_sync"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.save_config()
    
    def get_status(self):
        """קבלת סטטוס Git"""
        status = {
            "auto_sync_enabled": self.is_auto_sync_enabled(),
            "last_sync": self.config.get("git", {}).get("last_sync", "לא בוצע"),
            "repo_url": self.config.get("git", {}).get("repo_url", ""),
            "branch": self.config.get("git", {}).get("branch", "main")
        }
        
        # בדיקת סטטוס Git
        success, output = self.run_git_command(["git", "status", "--porcelain"])
        if success:
            status["has_changes"] = len(output.strip()) > 0
            status["changes_count"] = len([line for line in output.strip().split('\n') if line.strip()])
        else:
            status["has_changes"] = False
            status["changes_count"] = 0
            
        return status

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="סקריפט סינכרון נתונים אוטומטי")
    parser.add_argument("--force", action="store_true", help="הרץ סינכרון גם אם מושבת")
    parser.add_argument("--status", action="store_true", help="הצג סטטוס")
    
    args = parser.parse_args()
    
    sync_manager = GitSyncManager()
    
    if args.status:
        status = sync_manager.get_status()
        print("📊 סטטוס Git:")
        print(f"  סינכרון אוטומטי: {'מופעל' if status['auto_sync_enabled'] else 'מושבת'}")
        print(f"  סינכרון אחרון: {status['last_sync']}")
        print(f"  URL מאגר: {status['repo_url']}")
        print(f"  ענף: {status['branch']}")
        print(f"  שינויים: {status['changes_count']} קבצים")
        sys.exit(0)
    
    success = sync_manager.sync_data(force=args.force)
    if not success:
        sys.exit(1)
