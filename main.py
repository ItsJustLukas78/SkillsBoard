import time
import sys
import requests
import os
import dotenv
from openpyxl import Workbook

dotenv.load_dotenv()

API_KEY = os.getenv("API_KEY")

# Get the event id from the event code since the API does not use event code for other requests
def get_event_id(event_code):
    url = "https://www.robotevents.com/api/v2/events"
    new_headers = {"accept": "application/json", "Authorization": "Bearer " + API_KEY}
    new_params = {"sku": event_code}

    response = requests.get(url, params=new_params, headers=new_headers)

    if response.status_code == 200:
        try:
            response = response.json()["data"][0]["id"]
            return response
        except IndexError:
            print("Error: " + str(response.status_code) + " getting event id")
            return None
    else:
        print("Error: " + str(response.status_code) + " getting event id")

        # If request throttled, wait for the time specified in the Retry-After header + 1 second, or 5 in unspecified
        if response.status_code == 429:
            wait_time = int(response.headers.get("Retry-After", 5)) + 1
            print(f"Sleeping for {wait_time} seconds")
            time.sleep(wait_time)
            return get_event_id(event_code)
        else:
            return None


# Get a list of the teams from the event using the event id
def get_teams(event_id):
    url = "https://www.robotevents.com/api/v2/events/" + str(event_id) + "/teams"
    new_headers = {"accept": "application/json", "Authorization": "Bearer " + API_KEY}
    response = requests.get(url, headers=new_headers)

    teams = []

    if response.status_code != 200:
        print("Error: " + str(response.status_code) + " getting teams")

        # If request throttled, wait for the time specified in the Retry-After header + 1 second, or 5 in unspecified
        if response.status_code == 429:
            wait_time = int(response.headers.get("Retry-After", 5)) + 1
            print(f"Sleeping for {wait_time} seconds")
            time.sleep(wait_time)
            return get_teams(event_id)
        else:
            return None
    else:
        # Get teams from first page
        teams.extend(response.json()["data"])

        next_page_url = response.json()["meta"]["next_page_url"]

        while next_page_url is not None:
            response = requests.get(next_page_url, headers=new_headers)

            if response.status_code != 200:
                print("Error: " + str(response.status_code) + " getting teams")
                return None

            next_page_url = response.json()["meta"]["next_page_url"]
            teams.extend(response.json()["data"])

        return teams


def get_skills(team_id, season_id=173):
    url = "https://www.robotevents.com/api/v2/teams/" + str(team_id) + "/skills/"
    new_headers = {"accept": "application/json", "Authorization": "Bearer " + API_KEY}
    response = requests.get(url, headers=new_headers, params={"season": season_id})

    skills = []

    if response.status_code != 200:
        print("Error: " + str(response.status_code) + " getting skills")

        # If request throttled, wait for the time specified in the Retry-After header + 1 second, or 5 in unspecified
        if response.status_code == 429:
            wait_time = int(response.headers.get("Retry-After", 5)) + 1
            print(f"Sleeping for {wait_time} seconds")
            time.sleep(wait_time)
            return get_skills(team_id, season_id)
        else:
            return None
    else:
        # Get skills from first page
        skills.extend(response.json()["data"])

        next_page_url = response.json()["meta"]["next_page_url"]

        while next_page_url is not None:
            response = requests.get(next_page_url, headers=new_headers, params={"season": season_id})

            if response.status_code != 200:
                print("Error: " + str(response.status_code) + " getting awards")
                break

            next_page_url = response.json()["meta"]["next_page_url"]
            skills.extend(response.json()["data"])

    return skills


def get_data(teams):
    data = []

    for team in teams:
        print("Getting data for team " + str(team.get("number", "no number")) + "...")
        skills = get_skills(team.get("id", "no id"))
        print("Finished")

        if skills is None:
            continue

        try:
            driver_scores = [skill["score"] for skill in skills if skill["type"] == "driver"]
            programming_scores = [skill["score"] for skill in skills if skill["type"] == "programming"]

            best_driver_score = max(driver_scores)
            best_programming_score = max(programming_scores)

            best_sum = best_driver_score + best_programming_score

            data.append([team.get("number", "none"), team.get("team_name", "none"), driver_scores, programming_scores, best_driver_score, best_programming_score, best_sum])
        except Exception as e:
            print("Error: " + str(e))
            data.append([team.get("number", "none"), team.get("team_name", "none"), 0, 0, 0, 0, 0])
    
    return sorted(data, key=lambda x: x[6], reverse=True)

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python3 skills_ranked.py <event_code> <output_file>")
        sys.exit(1)
    elif len(sys.argv) == 3:
        event_id = get_event_id(sys.argv[1])
        if event_id is None:
            print("Error: could not get event from that code")
            sys.exit(1)

    event_id = get_event_id(sys.argv[1])
    teams = get_teams(event_id)
    team_to_awards = get_data(teams)

    workbook = Workbook()
    sheet = workbook.active

    sheet["A1"] = "Team number"
    sheet["B1"] = "Team name"
    sheet["C1"] = "Driver Scores"
    sheet["D1"] = "Prog Scores"
    sheet["D1"] = "Prog Scores"
    sheet["E1"] = "Best Driver Score"
    sheet["F1"] = "Best Prog Score"
    sheet["G1"] = "Best Sum"

    for x in range(0, len(team_to_awards)):
        sheet.append((str(team_data) for team_data in team_to_awards[x]))

    try:
        workbook.save(filename=f"{sys.argv[2]}.xlsx")
    except Exception as e:
        # Output file name didn't work, use the event code
        print("Error: " + str(e))
        workbook.save(filename=f"{sys.argv[1]}_skills.xlsx")