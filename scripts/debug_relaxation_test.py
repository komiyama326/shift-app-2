import datetime
import random

from core_engine import SettingsManager, generate_calendar_with_holidays, ShiftScheduler


def run_case_fixed_for_a(fixed_days: int, seed: int = 0):
    random.seed(seed)

    sm = SettingsManager()
    # スタッフ8人を作成（色は仮）
    colors = ['#FF0000','#00FF00','#0000FF','#FFA500','#800080','#008080','#808000','#000000']
    names = ['A','B','C','D','E','F','G','H']
    for i, n in enumerate(names):
        sm.staff_manager.add_or_update_staff(
            __import__('core_engine').Staff(n, colors[i % len(colors)], set(), True)
        )

    # 年月
    today = datetime.date.today()
    year = today.year
    month = today.month
    cal = generate_calendar_with_holidays(year, month)

    # A に固定日をランダムに fixed_days 日設定
    all_days = [d['date'] for d in cal]
    chosen = sorted(random.sample(all_days, min(fixed_days, len(all_days))))
    a = sm.staff_manager.get_staff_by_name('A')
    manual_fixed = {}
    for d in chosen:
        manual_fixed.setdefault(d, []).append(a)

    # スケジューラ実行
    scheduler = ShiftScheduler(sm.staff_manager, cal)
    res = scheduler.solve(
        shifts_per_day={'min':1,'max':1},
        min_interval=5,
        max_consecutive_days=5,
        max_solutions=1,
        no_shift_dates=[],
        manual_fixed_shifts=manual_fixed,
        rule_based_fixed_shifts=[],
        vacations={},
        rule_based_vacations=[],
        fairness_group=set(['月','火','水','木','金','土','日']),
        fairness_tolerance=1,
        disperse_duties=True,
        past_schedules={},
        fairness_as_hard=True,
        fallback_soft_on_infeasible=False
    )

    if isinstance(res, str):
        print('INFEASIBLE\n' + res)
    else:
        sol = res[0]
        print('FEASIBLE')
        print('relaxations:', sol.get('relaxations'))
        print('counts:', sol.get('counts'))


if __name__ == '__main__':
    print('== A fixed 10 days ==')
    run_case_fixed_for_a(10, seed=42)
    print('\n== A fixed 5 days ==')
    run_case_fixed_for_a(5, seed=43)

