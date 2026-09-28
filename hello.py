import json
import openpyxl as xl
from datetime import datetime

# ---------------- CUSTOM EXCEPTION ----------------
class InvalidRecordError(Exception):
    pass

# ---------------- DAILY RECORD CLASS ----------------
class DailyRecord:
    def __init__(self, date, sleep, fitness, study, coding, classes,
                 attended, other, total, free, feeling, satisfaction, energy):
        self.date = date
        self.sleep = sleep
        self.fitness = fitness
        self.study = study
        self.coding = coding
        self.classes = classes
        self.attended = attended
        self.other = other
        self.total = total
        self.free = free
        self.feeling = feeling
        self.satisfaction = satisfaction
        self.energy = energy

    def to_dict(self):
        return self.__dict__


# ---------------- VALIDATION ----------------
FEELINGS = ("Excellent", "Good", "Neutral", "Low", "Stressed")
SATISFACTIONS = ("Very Satisfied", "Satisfied", "Neutral",
                 "Unsatisfied", "Very Unsatisfied")
ENERGIES = ("High", "Medium", "Low")

def validate(r):
    numbers = [r.sleep, r.fitness, r.study, r.coding, r.classes,
               r.attended, r.other, r.total, r.free]

    if any(not isinstance(x, (int, float)) or x < 0 for x in numbers):
        raise InvalidRecordError(f"Invalid numeric value on {r.date}")

    if r.feeling not in FEELINGS:
        raise InvalidRecordError(f"Invalid feeling on {r.date}")

    if r.satisfaction not in SATISFACTIONS:
        raise InvalidRecordError(f"Invalid satisfaction on {r.date}")

    if r.energy not in ENERGIES:
        raise InvalidRecordError(f"Invalid energy on {r.date}")


# ---------------- READ EXCEL ----------------
def load_data(filename):
    try:
        wb = xl.load_workbook(filename, data_only=True)
        ws = wb["Daily Log"]
    except FileNotFoundError:
        print("File not found:", filename)
        return [],[]
    except Exception as e:
        print("Error:", e)
        return [],[]

    records = []
    invalid = []

    for row in ws.iter_rows(min_row=6, values_only=True):
        if not row[0]:
            continue

        try:
            r = DailyRecord(
                row[0], row[1], row[2], row[3], row[4], row[5],
                row[6], row[7], row[8], row[9], row[10], row[11], row[12]
            )
            validate(r)
            records.append(r)

        except (InvalidRecordError, TypeError, ValueError) as e:
            invalid.append((row[0], str(e)))

    wb.close()
    return records, invalid


# ---------------- STATISTICS ----------------
def statistics(records):
    n = len(records)

    if n == 0:
        print("No valid records.")
        return

    fields = {
        "Sleep": "sleep",
        "Fitness": "fitness",
        "Study": "study",
        "Coding": "coding",
        "Class": "classes",
        "Other Activity": "other",
        "Free Time": "free"
    }

    print("\n========== AVERAGE ACTIVITY ==========")
    print("Valid Days:", n)

    for name, field in fields.items():
        values = [getattr(r, field) for r in records]
        print(f"Average {name}: {sum(values)/n:.2f} min")

    print("\n========== MIN / MAX ==========")
    coding = [r.coding for r in records]
    study = [r.study for r in records]

    print("Minimum Coding:", min(coding), "min")
    print("Maximum Coding:", max(coding), "min")
    print("Minimum Study:", min(study), "min")
    print("Maximum Study:", max(study), "min")


# ---------------- INDEX CALCULATIONS ----------------
def indexes(records):
    n = len(records)

    tpi = sum(r.coding for r in records) / n
    aai = sum(r.study + r.classes for r in records) / n
    phai = sum(r.fitness for r in records) / n
    sri = sum(r.sleep for r in records) / n
    abi = sum(r.free for r in records) / n
    tui = sum(r.total for r in records) / n

    score = {
        "Excellent": 5, "Good": 4, "Neutral": 3,
        "Low": 2, "Stressed": 1,
        "Very Satisfied": 5, "Satisfied": 4,
        "Unsatisfied": 2, "Very Unsatisfied": 1,
        "High": 3, "Medium": 2
    }

    ei = sum(
        score[r.feeling] +
        score[r.satisfaction] +
        score[r.energy]
        for r in records
    ) / (n * 3)

    dci = n / 36 * 100

    pai = (
        .15*tpi + .20*aai + .15*phai +
        .20*sri + .15*tui + .10*ei + .05*dci
    )

    print("\n========== INDEX VALUES ==========")
    print(f"Tech Productivity      : {tpi:.2f}")
    print(f"Academic Activity      : {aai:.2f}")
    print(f"Physical Activity      : {phai:.2f}")
    print(f"Sleep Recovery         : {sri:.2f}")
    print(f"Activity Balance       : {abi:.2f}")
    print(f"Time Utilization       : {tui:.2f}")
    print(f"Experience Index       : {ei:.2f}")
    print(f"Data Continuity        : {dci:.2f}%")
    print(f"Personal Activity      : {pai:.2f}")


# ---------------- CATEGORY SUMMARY ----------------
def category_summary(records):
    feelings = {}
    satisfactions = {}
    energies = set()

    for r in records:
        feelings[r.feeling] = feelings.get(r.feeling, 0) + 1
        satisfactions[r.satisfaction] = satisfactions.get(
            r.satisfaction, 0) + 1
        energies.add(r.energy)

    print("\n========== CATEGORY SUMMARY ==========")
    print("Feelings:", feelings)
    print("Satisfaction:", satisfactions)
    print("Energy Levels:", energies)


# ---------------- QUERY ----------------
def query(records):
    print("\n1. Satisfaction")
    print("2. Feeling")
    print("3. Energy")
    print("4. High Coding")
    choice = input("Enter choice: ")

    if choice == "1":
        value = input("Enter satisfaction: ")
        result = [r for r in records if r.satisfaction == value]

    elif choice == "2":
        value = input("Enter feeling: ")
        result = [r for r in records if r.feeling == value]

    elif choice == "3":
        value = input("Enter energy: ")
        result = [r for r in records if r.energy == value]

    elif choice == "4":
        result = [r for r in records if r.coding > 120]

    else:
        print("Invalid choice.")
        return

    for r in result:
        print(r.date, "|", r.feeling, "|",
              r.satisfaction, "| Coding:", r.coding)


# ---------------- JSON ----------------
def save_json(records):
    data = [r.to_dict() for r in records]

    with open("12617630_cleaned.json", "w") as f:
        json.dump(data, f, indent=4, default=str)

    print("Cleaned data saved to 12617630_cleaned.json")


# ---------------- MAIN MENU ----------------
def main():
    records, invalid = load_data("12617630.xlsx")

    if not records:
        print("No valid records found.")
        return

    while True:
        print("\n========== DAILY ROUTINE ANALYZER ==========")
        print("1. Statistics")
        print("2. Index Values")
        print("3. Category Summary")
        print("4. Query Records")
        print("5. Save Cleaned JSON")
        print("6. Show Invalid Records")
        print("7. Exit")

        choice = input("Enter choice: ")

        if choice == "1":
            statistics(records)

        elif choice == "2":
            indexes(records)

        elif choice == "3":
            category_summary(records)

        elif choice == "4":
            query(records)

        elif choice == "5":
            save_json(records)

        elif choice == "6":
            print("\nInvalid Records:", len(invalid))
            for item in invalid:
                print(item)

        elif choice == "7":
            print("Program ended.")
            break

        else:
            print("Invalid choice.")
main()