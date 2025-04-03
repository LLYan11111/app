from pymongo import MongoClient
from datetime import datetime
import os
import sys
import json
import logging
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

def get_config_path():
    """Get the configuration file path, handling both packaged and development environments"""
    if getattr(sys, 'frozen', False):
        # Path for packaged executable
        base_dir = os.path.dirname(sys.executable)
    else:
        # Development environment
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    
    return os.path.join(base_dir, 'database', 'mongo_config.json')

def get_database():
    """Get MongoDB connection from configuration file"""
    from pymongo import MongoClient
    
    try:
        config_path = get_config_path()
        if not os.path.exists(config_path):
            raise FileNotFoundError(f"MongoDB configuration file does not exist: {config_path}")
            
        with open(config_path, 'r') as f:
            config = json.load(f)
            
        # Check if either database_name or database key exists
        if 'database_name' in config:
            db_name = config['database_name']
        elif 'database' in config:
            db_name = config['database']
        else:
            raise KeyError("MongoDB configuration file is missing database name setting (either 'database_name' or 'database')")
        
        # Check connection string
        if 'connection_string' not in config:
            raise KeyError("MongoDB configuration file is missing 'connection_string' setting")
            
        connection_string = config['connection_string']
        
        # Display connection info (optional)
        if getattr(sys, 'frozen', False):
            print(f"Connecting to MongoDB: {connection_string} ({db_name})")
            
        client = MongoClient(connection_string)
        return client[db_name]
    except Exception as e:
        import logging
        logging.error(f"MongoDB connection error: {str(e)}")
        print(f"Database error: {str(e)}")
        if getattr(sys, 'frozen', False):
            input("Press Enter to exit...")
        sys.exit(1)

def init_database():
    """Initialize MongoDB collections and indexes"""
    try:
        db = get_database()
        if db is None:
            logging.warning("Unable to initialize MongoDB, will use local storage")
            return False
            
        # Create collections
        collections = ['users', 'activities', 'idle_times', 'afk']
        for collection in collections:
            if collection not in db.list_collection_names():
                db.create_collection(collection)
        
        # Create indexes
        db.users.create_index('username', unique=True)
        db.activities.create_index([('date', 1)])
        db.idle_times.create_index([('user_name', 1), ('date', 1)])
        db.afk.create_index([('start', 1)])
        db.afk.create_index([('type', 1)])
        
        logging.info("MongoDB initialization successful")
        return True
        
    except Exception as e:
        logging.error(f"Error initializing MongoDB: {e}")
        print(f"Error initializing MongoDB: {e}")
        return False