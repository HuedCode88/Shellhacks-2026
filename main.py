from database import initialize_database, load_all_projects
from server import serve_map

def main():
    initialize_database()
    projects = load_all_projects()

    print(f"Loaded {len(projects)} geocoded projects.")
    serve_map()

if __name__ == "__main__":
    main()