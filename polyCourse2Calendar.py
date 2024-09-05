import csv
from collections import defaultdict
from typing import Dict, Optional, List
from ics import Calendar, Event
from datetime import datetime, timedelta
import pytz

# Set the Montreal timezone
montreal_tz = pytz.timezone('America/Toronto')

# Mapping of full French weekday names to abbreviations
FULL_DAY_TO_ABBREVIATION = {
    "LUNDI": "LUN",
    "MARDI": "MAR",
    "MERCREDI": "MER",
    "JEUDI": "JEU",
    "VENDREDI": "VEN",
    "SAMEDI": "SAM",
    "DIMANCHE": "DIM"
}

class Group:
    def __init__(self, number: int, week_day: str, room: str, course_length: int, frequency: str, hour: str):
        self.number = number
        self.week_day = week_day
        self.room = room
        self.course_length = course_length 
        self.frequency = frequency  
        self.hour = ','.join(sorted(hour.split(',')))

    def __repr__(self):
        return (f"Group(number={self.number}, week_day={self.week_day}, room={self.room}, "
                f"course_length={self.course_length}, frequency={self.frequency}, hour={self.hour})")


class Course:
    def __init__(self, sigle: str, name: str, nb_credit: int):
        self.sigle = sigle
        self.name = name
        self.nb_credit = nb_credit
        self.theo_groups: Dict[int, List[Group]] = defaultdict(list)
        self.lab_groups: Dict[int, List[Group]] = defaultdict(list)

    def add_group(self, group_type: str, group: Group):
        if group_type.lower() == "c":
            self.theo_groups[group.number].append(group)
        elif group_type.lower() == "l":
            self.lab_groups[group.number].append(group)
        self.merge_groups() 

    def merge_groups(self):
        """Automatically merge groups with the same type, room, day, hour, and frequency."""
        self._merge_group_type(self.theo_groups)
        self._merge_group_type(self.lab_groups)

    def _merge_group_type(self, groups: Dict[int, List[Group]]):
        """Helper function to merge groups of a particular type that occur on the same day, room, hour, and frequency."""
        for number, group_list in groups.items():
            merged_groups = defaultdict(int) 
            hour_map = defaultdict(list) 
            
            for group in group_list:
                key = (group.week_day, group.room, group.frequency)
                merged_groups[key] += group.course_length
                hour_map[key].append(group.hour)
            
            groups[number] = [
                Group(number, day, room, duration, frequency, ','.join(hours))
                for (day, room, frequency), duration in merged_groups.items()
                for hours in [hour_map[(day, room, frequency)]]
            ]

    def get_all_groups(self) -> List[Group]:
        """Return all groups (theo and lab) for the course."""
        return [group for groups in self.theo_groups.values() for group in groups] + \
               [group for groups in self.lab_groups.values() for group in groups]

    def get_specific_group(self, group_type: str, group_number: int) -> Optional[List[Group]]:
        """
        Retrieve a specific group based on the type and group number.
        
        Args:
            group_type (str): The type of the group ("c" for theoretical, "l" for lab).
            group_number (int): The number of the group.
        
        Returns:
            Optional[List[Group]]: The list of groups matching the criteria, or None if not found.
        """
        if group_type.lower() == "c":
            return self.theo_groups.get(group_number)
        elif group_type.lower() == "l":
            return self.lab_groups.get(group_number)
        return None

    def __repr__(self):
        return (f"Course(sigle={self.sigle}, name={self.name}, nb_credit={self.nb_credit}, "
                f"theo_groups={self.theo_groups}, lab_groups={self.lab_groups})")


class Courses:
    def __init__(self):
        self.courses: Dict[str, Course] = {}

    def read_csv_files(self, horsages_path: str):
        with open(horsages_path, newline='', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile, delimiter=';')
            next(reader)  

            for row in reader:
                if len(row) < 15:
                    continue

                _, sigle, number, nb_credit, _, _, room, period_type, _, week_nb, course_type, name, _, week_day, hour = row

                if not sigle or not number or not nb_credit or not room or not period_type or not course_type or not name or not week_day or not hour:
                    continue

                try:
                    number = int(number)
                    nb_credit = int(float(nb_credit.replace(',', '.')))
                    course_length = 1 

                    if week_nb == "I":
                        frequency = "B1"
                    elif week_nb == "P":
                        frequency = "B2"
                    else:
                        frequency = "every_week" 
                except ValueError:
                    continue

                if sigle not in self.courses:
                    self.courses[sigle] = Course(sigle, name, nb_credit)

                group = Group(number, week_day, room, course_length, frequency, hour)

                self.courses[sigle].add_group(period_type.lower(), group)
    
    def get_course(self, sigle: str) -> Optional[Course]:
        """Retrieve a course by its sigle (course code)."""
        return self.courses.get(sigle)

    def __repr__(self):
        return f"Courses({self.courses})"

    def read_alternance_csv(self, alternance_path: str) -> Dict[datetime, Dict[str, str]]:
        week_map = {}
        with open(alternance_path, newline='', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile, delimiter=',')
            next(reader) 
            for row in reader:
                if len(row) < 3:
                    continue
                date_str, day_name, week_type = row
                try:
                    date = datetime.strptime(date_str, '%Y-%m-%d')
                    week_map[date] = {
                        "day_name": day_name.strip().upper(),
                        "week_type": week_type.strip().upper()
                    }
                except ValueError:
                    continue
        return week_map

def generate_ics_file(courses_manager: 'Courses', courses_to_convert: Dict[str, Dict[str, List[int]]], filename: str, alternance_map: Dict[datetime, Dict[str, str]]) -> Calendar:
    calendar = Calendar()

    semester_dates = sorted(alternance_map.keys())
    semester_start_date = semester_dates[0]
    semester_end_date = semester_dates[-1]

    for course_code, group_selection in courses_to_convert.items():
        course = courses_manager.get_course(course_code)
        if course:
            for group_type, group_numbers in group_selection.items():
                for group_number in group_numbers:
                    groups = course.get_specific_group(group_type, group_number)
                    if groups:
                        for group in groups:
                            frequency = group.frequency.lower()

                            current_date = semester_start_date
                            while current_date <= semester_end_date:
                                alternance_info = alternance_map.get(current_date)
                                
                                if alternance_info:
                                    full_day_name = alternance_info["day_name"]
                                    week_type = alternance_info["week_type"]
                                    day_abbr = FULL_DAY_TO_ABBREVIATION.get(full_day_name, full_day_name)
                                    
                                    # Special case: Treat October 1st as a Monday (LUN) School Calendar
                                    if current_date == datetime(2024, 10, 1):
                                        day_abbr = "LUN"

                                else:
                                    current_date += timedelta(days=1)
                                    continue
                                
                                if day_abbr == group.week_day.upper() and (
                                    frequency == "every_week" or
                                    (frequency == "b1" and week_type == "B1") or
                                    (frequency == "b2" and week_type == "B2")
                                ):
                                    # Convert the start time from "HHMM" format to datetime
                                    start_time_str = group.hour.split(',')[0]
                                    start_time = datetime.strptime(start_time_str, "%H%M")
                                    
                                    event_date_time = current_date.replace(hour=start_time.hour, minute=start_time.minute)
                                    event_date_time = montreal_tz.localize(event_date_time)

                                    total_hours = len(group.hour.split(','))

                                    event = Event()
                                    event.name = f"{course.sigle} - {course.name} - Groupe {group.number} ({'COURS' if group_type == 'c' else 'LAB'})"
                                    event.begin = event_date_time
                                    event.duration = timedelta(hours=total_hours)
                                    event.location = group.room

                                    event.alarms = [
                                        {"action": "display", "trigger": timedelta(minutes=-30)}
                                    ]

                                    calendar.events.add(event)

                                current_date += timedelta(days=1)

    with open(filename, 'w', encoding='utf-8') as f:
        f.writelines(calendar)

horsages_path = 'horsage.csv'
alternance_path = 'alternance.csv'
courses_manager = Courses()
courses_manager.read_csv_files(horsages_path)

alternance_map = courses_manager.read_alternance_csv(alternance_path)

courses_to_convert = {
    "LOG1810": {"c": [2], "l": [4]},
    "INF1015": {"c": [1], "l": [1]},
}

generate_ics_file(courses_manager, courses_to_convert, "schedule.ics", alternance_map)
