import os
import sys
import datetime

# Add repo root to sys.path so we can `import core_engine`
sys.path.append(os.path.dirname(os.path.dirname(__file__)))
from core_engine import Staff, StaffManager, ShiftScheduler, generate_calendar_with_holidays


def main():
    sm = StaffManager()
    A = Staff('A', '#ff0000', impossible_weekdays={'火'})
    B = Staff('B', '#00ff00')
    sm.add_or_update_staff(A)
    sm.add_or_update_staff(B)

    year, month = 2024, 9
    cal = generate_calendar_with_holidays(year, month)

    tuesdays = [d['date'] for d in cal if d['date'].weekday() == 1]
    second_tuesday = tuesdays[1]

    manual_vacations = {'A': [second_tuesday]}

    manual_fixed = {second_tuesday: [A]}
    first = second_tuesday + datetime.timedelta(days=1)
    second = second_tuesday + datetime.timedelta(days=2)
    manual_fixed[first] = manual_fixed.get(first, []) + [A]
    manual_fixed[second] = manual_fixed.get(second, []) + [A]

    scheduler = ShiftScheduler(sm, cal, ignore_rules_on_holidays=True)
    solutions = scheduler.solve(
        shifts_per_day={'min': 1, 'max': 1},
        min_interval=2,
        max_consecutive_days=5,
        max_solutions=1,
        last_month_end_dates={},
        prev_month_consecutive_days={},
        last_week_assignments={},
        avoid_consecutive_same_weekday=False,
        no_shift_dates=[],
        manual_fixed_shifts=manual_fixed,
        rule_based_fixed_shifts=[],
        vacations=manual_vacations,
        rule_based_vacations=[],
        fairness_group=set(),
        total_adjustments={},
        fairness_adjustments={},
        disperse_duties=False,
        past_schedules={}
    )

    if isinstance(solutions, str):
        print('ERROR:', solutions)
        return 1

    sol = solutions[0]
    schedule = sol["schedule"]
    assigned = {date: [s.name for s in staff_list] for date, staff_list in schedule.items()}
    print('Second Tuesday:', second_tuesday, 'Assigned:', assigned.get(second_tuesday))
    print('Next Day:', first, 'Assigned:', assigned.get(first))
    print('Next Next Day:', second, 'Assigned:', assigned.get(second))
    ok_vacation = ('A' not in assigned.get(second_tuesday, []))
    ok_fixed_next = ('A' in assigned.get(first, []))
    ok_fixed_next2 = ('A' in assigned.get(second, []))
    print('Checks:', {'vacation_wins': ok_vacation, 'fixed_next': ok_fixed_next, 'fixed_next2': ok_fixed_next2})
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
