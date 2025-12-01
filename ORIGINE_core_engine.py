# core_engine.py

import calendar
import datetime
import holidays
from dataclasses import dataclass
import json
import time
import os

from ortools.sat.python import cp_model
from collections import defaultdict

weekdays_jp = ("月", "火", "水", "木", "金", "土", "日")

# ★★★ 削除: RuleBasedImpossibleDay クラス ★★★

class Staff:
    # ★★★ 修正: rule_based_impossible_days を削除 ★★★
    def __init__(self, name: str, color_code: str, impossible_weekdays: set[str] = None, is_active: bool = True):
        if not isinstance(name, str) or not name:
            raise ValueError("スタッフ名は必須です。")
        if not isinstance(color_code, str) or not color_code.startswith('#'):
            raise ValueError("カラーコードは'#'で始まる文字列である必要があります。")
        self.name = name
        self.color_code = color_code
        self.impossible_weekdays = set(impossible_weekdays) if impossible_weekdays is not None else set()
        self.is_active = is_active

    def is_available(self, weekday: str) -> bool:
        return weekday not in self.impossible_weekdays
        
    def __repr__(self):
        return (f"Staff(名前: {self.name}, 色: {self.color_code}, 稼働中: {self.is_active}, "
                f"不可曜日: {self.impossible_weekdays or 'なし'})")

class StaffManager:
    def __init__(self):
        self.staff_map: dict[str, Staff] = {}
    def add_or_update_staff(self, staff: Staff):
        if staff.name in self.staff_map:
            print(f"スタッフ '{staff.name}' の情報を更新します。")
        else:
            print(f"スタッフ '{staff.name}' を追加します。")
        self.staff_map[staff.name] = staff
    def remove_staff_by_name(self, name: str) -> bool:
        if name in self.staff_map:
            del self.staff_map[name]
            print(f"スタッフ '{name}' を削除しました。")
            return True
        print(f"警告: スタッフ '{name}' が見つかりません。")
        return False
    def get_all_staff(self) -> list[Staff]:
        return list(self.staff_map.values())
    
    def get_active_staff(self) -> list[Staff]:
        return [staff for staff in self.staff_map.values() if staff.is_active]

    def __len__(self):
        return len(self.staff_map)
    def get_staff_by_name(self, name: str) -> Staff | None:
        return self.staff_map.get(name)

@dataclass(frozen=True)
class RuleBasedFixedShift:
    week_number: int
    weekday_index: int
    staff: Staff    

@dataclass(frozen=True)
class RuleBasedVacation:
    week_number: int
    weekday_index: int
    staff_name: str

class ShiftScheduler:
    def __init__(self, staff_manager: StaffManager, calendar_data: list[dict], ignore_rules_on_holidays: bool = False):
        self.staff_manager = staff_manager
        self.calendar_data = calendar_data
        self.all_staff = self.staff_manager.get_active_staff()
        self.ignore_rules_on_holidays = ignore_rules_on_holidays
        self.jp_holidays = holidays.JP(years=self.calendar_data[0]['date'].year)

    def solve(self, 
              shifts_per_day: int | dict[str, int] = 1,
              min_interval: int = 2, 
              max_consecutive_days: int = 5,
              max_solutions: int = 1,
              last_month_end_dates: dict = None,
              prev_month_consecutive_days: dict = None,
              last_week_assignments: dict = None,
              avoid_consecutive_same_weekday: bool = False,
              no_shift_dates: list = None,
              manual_fixed_shifts: dict = None,
              rule_based_fixed_shifts: list = None,
              vacations: dict = None,
              rule_based_vacations: list = None,
              fairness_group: set = None,
              total_adjustments: dict = None,
              fairness_adjustments: dict = None,
              fairness_tolerance: int = 1,
              disperse_duties: bool = True,
              past_schedules: dict = None,
              **kwargs
              ) -> list[dict] | str:

        staff_list = self.all_staff
        if not staff_list: return []
        
        found_solutions = []
        for i in range(max_solutions):
            model = cp_model.CpModel()
            self.constraint_tags = {}
            shifts = self._define_variables(model, staff_list, self.calendar_data)
            
            self._add_hard_constraints(model, shifts, staff_list, self.calendar_data,
                                       no_shift_dates, shifts_per_day,
                                       rule_based_vacations, vacations,
                                       min_interval, max_consecutive_days,
                                       last_month_end_dates, prev_month_consecutive_days,
                                       fairness_group, avoid_consecutive_same_weekday,
                                       last_week_assignments)
            
            _ , fixed_shift_penalty = self._add_soft_constraints(
                model, shifts, staff_list, self.calendar_data,
                rule_based_fixed_shifts, manual_fixed_shifts
            )

            dispersion_penalty = 0
            if disperse_duties and fairness_group:
                dispersion_penalty = self._add_dispersion_penalty(
                    model, shifts, staff_list, self.calendar_data,
                    fairness_group, past_schedules
                )

            self._add_fairness_objective(model, shifts, staff_list, self.calendar_data,
                                         total_adjustments, fairness_adjustments, 
                                         fairness_tolerance, fairness_group,
                                         fixed_shift_penalty,
                                         dispersion_penalty)

            for sol in found_solutions:
                self._add_solution_prohibition_constraint(model, shifts, sol)
            
            solver = cp_model.CpSolver()
            print(f"CP-SATソルバーによるシフト生成を開始します... ({len(found_solutions) + 1}件目の探索)")
            status = solver.Solve(model)

            if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
                print("解を発見しました。")
                solution = self._create_solution_from_solver(solver, shifts, staff_list, self.calendar_data, fairness_group)
                found_solutions.append(solution)
            else:
                if status == cp_model.INFEASIBLE:
                    print("解が見つかりませんでした。制約の衝突を分析します。")
                    error_report = self._analyze_infeasibility(solver)
                    return error_report
                else:
                    print("これ以上、新しい解は見つかりませんでした。")
                    break

        print(f"探索完了。合計 {len(found_solutions)} 件の異なる解が見つかりました。")
        return found_solutions
    
    def _define_variables(self, model, staff_list, day_list):
        shifts = {}
        for s in range(len(staff_list)):
            for d in range(len(day_list)):
                shifts[(s, d)] = model.NewBoolVar(f'shift_s{s}_d{d}')
        return shifts

    def _add_hard_constraints(self, model, shifts, staff_list, day_list,
                              no_shift_dates, shifts_per_day_config,
                              rule_based_vacations, vacations,
                              min_interval, max_consecutive_days,
                              last_month_end_dates, prev_month_consecutive_days,
                              fairness_group, avoid_consecutive_same_weekday,
                              last_week_assignments):
        
        num_staff = len(staff_list)
        year = day_list[0]['date'].year
        month = day_list[0]['date'].month
        first_day_of_month = day_list[0]['date']

        for d, day_info in enumerate(day_list):
            date_obj = day_info['date']
            tag = f"{date_obj.day}日の必要人数"
            if no_shift_dates and date_obj in no_shift_dates:
                 c = model.Add(sum(shifts[(s, d)] for s in range(num_staff)) == 0)
                 self.constraint_tags[c.Index()] = tag + "（不要日）"
                 continue
            
            min_needed, max_needed = self._get_shift_range_for_day(day_info, shifts_per_day_config)
            
            c = model.AddLinearConstraint(sum(shifts[(s, d)] for s in range(num_staff)), min_needed, max_needed)
            self.constraint_tags[c.Index()] = tag

        generated_vacations = self._generate_vacations_from_rules(rule_based_vacations, year, month)
        all_vacations = vacations if vacations else {}

        # ★★★ 修正: _generate_all_staff_impossible_dates の呼び出しを削除 ★★★
        all_impossible_dates_map = {}
        for staff in staff_list:
            impossible_set = set()
            impossible_set.update(generated_vacations.get(staff.name, set()))
            impossible_set.update(all_vacations.get(staff.name, set()))
            all_impossible_dates_map[staff.name] = impossible_set

        for s, staff in enumerate(staff_list):
            staff_impossible_dates = all_impossible_dates_map.get(staff.name, set())
            for d, day_info in enumerate(day_list):
                date_obj = day_info['date']
                if date_obj in staff_impossible_dates:
                    c = model.Add(shifts[(s, d)] == 0)
                    self.constraint_tags[c.Index()] = f"{staff.name}の{date_obj.day}日の休暇/不可日"
                    continue
                is_holiday_and_ignored = self.ignore_rules_on_holidays and date_obj in self.jp_holidays
                if not staff.is_available(day_info['weekday']):
                    if not is_holiday_and_ignored:
                        c = model.Add(shifts[(s, d)] == 0)
                        self.constraint_tags[c.Index()] = f"{staff.name}の{day_info['weekday']}曜日の不可日"
            
            if last_month_end_dates and staff.name in last_month_end_dates:
                last_worked_date = last_month_end_dates[staff.name]
                days_since_last = (first_day_of_month - last_worked_date).days
                
                days_to_forbid = min_interval - days_since_last + 1
                for d in range(days_to_forbid):
                    if d < len(day_list):
                        c = model.Add(shifts[(s, d)] == 0)
                        tag = f"{staff.name}の{day_list[d]['date'].day}日の勤務不可（前月からの間隔）"
                        self.constraint_tags[c.Index()] = tag
            
            for d in range(len(day_list) - min_interval - 1):
                lit_d = shifts[(s, d)]
                lit_d1_not = shifts[(s, d + 1)].Not()
                
                for i in range(1, min_interval):
                    c = model.AddImplication(shifts[(s, d + 1 + i)], lit_d.Not()).OnlyEnforceIf([lit_d, lit_d1_not])
                    self.constraint_tags[c.Index()] = f"{staff.name}の{day_list[d]['date'].day}日からの休み間隔"

            for d in range(len(day_list) - max_consecutive_days):
                window = [shifts[(s, i)] for i in range(d, d + max_consecutive_days + 1)]
                c = model.Add(sum(window) <= max_consecutive_days)
                self.constraint_tags[c.Index()] = f"{staff.name}の{day_list[d]['date'].day}日からの最大連勤"
            
            if prev_month_consecutive_days and staff.name in prev_month_consecutive_days:
                consecutive = prev_month_consecutive_days[staff.name]
                if consecutive > 0:
                    remaining_days = max_consecutive_days - consecutive
                    if remaining_days < max_consecutive_days:
                        window = [shifts[(s, d)] for d in range(remaining_days + 1)]
                        c = model.Add(sum(window) <= remaining_days)
                        self.constraint_tags[c.Index()] = f"{staff.name}の月初の連勤制限"

    def _add_soft_constraints(self, model, shifts, staff_list, day_list,
                              rule_based_fixed_shifts, manual_fixed_shifts):
        year = day_list[0]['date'].year
        month = day_list[0]['date'].month
        
        penalty_literals = []
        penalty_cost = model.NewIntVar(0, 0, 'empty_penalty')

        generated_fixed_shifts = self._generate_fixed_shifts_from_rules(rule_based_fixed_shifts, year, month)
        all_fixed_shifts = {**generated_fixed_shifts, **(manual_fixed_shifts if manual_fixed_shifts else {})}
        
        for date_obj, staff_obj_list in all_fixed_shifts.items():
            try:
                d = [i for i, day in enumerate(day_list) if day['date'] == date_obj][0]
                for staff_obj in staff_obj_list:
                    if staff_obj not in staff_list: continue
                    s = staff_list.index(staff_obj)
                    
                    lit = model.NewBoolVar(f"fixed_penalty_s{s}_d{d}")
                    model.Add(shifts[(s, d)] == 0).OnlyEnforceIf(lit)
                    model.Add(shifts[(s, d)] == 1).OnlyEnforceIf(lit.Not())
                    
                    penalty_literals.append(lit)
                    tag = f"{staff_obj.name}の{date_obj.day}日の固定シフト"
                    self.constraint_tags[lit] = tag

            except (ValueError, IndexError):
                print(f"警告: 固定シフトの日付 {date_obj} がカレンダーに見つかりません。")
        
        if penalty_literals:
            penalty_cost = model.NewIntVar(0, len(penalty_literals), 'penalty_cost')
            model.Add(penalty_cost == sum(penalty_literals))
        
        return penalty_literals, penalty_cost

    def _add_fairness_objective(self, model, shifts, staff_list, day_list,
                                total_adjustments, fairness_adjustments, 
                                fairness_tolerance, fairness_group,
                                fixed_shift_penalty,
                                dispersion_penalty):
        num_staff = len(staff_list)
        num_days = len(day_list)

        total_penalty = model.NewIntVar(0, 1000000, 'total_penalty')
        model.Add(total_penalty == fixed_shift_penalty + dispersion_penalty)

        if num_staff > 1:
            total_shifts = [model.NewIntVar(0, num_days, f'total_s{s}') for s in range(num_staff)]
            adj_total = [model.NewIntVar(-num_days, num_days, f'adj_total_s{s}') for s in range(num_staff)]
            for s, staff in enumerate(staff_list):
                model.Add(total_shifts[s] == sum(shifts[(s, d)] for d in range(num_days)))
                adj = total_adjustments.get(staff.name, 0) if total_adjustments else 0
                model.Add(adj_total[s] == total_shifts[s] - adj)
            
            min_total = model.NewIntVar(-num_days, num_days, 'min_total')
            max_total = model.NewIntVar(0, num_days, 'max_total')
            model.AddMinEquality(min_total, adj_total)
            model.AddMaxEquality(max_total, adj_total)
            total_diff = model.NewIntVar(0, num_days, 'total_diff')
            model.Add(total_diff == max_total - min_total)
            
            c1 = model.Add(total_diff <= fairness_tolerance)
            self.constraint_tags[c1.Index()] = f"総回数の公平性 (許容差: {fairness_tolerance}回)"

            if fairness_group:
                special_day_indices = [
                    d for d, day_info in enumerate(day_list)
                    if ('祝' in fairness_group and day_info.get('is_national_holiday', False)) or \
                       (day_info['weekday'] in fairness_group)
                ]
                
                if special_day_indices:
                    fair_shifts = [model.NewIntVar(0, len(special_day_indices), f'fair_s{s}') for s in range(num_staff)]
                    adj_fair = [model.NewIntVar(-num_days, num_days, f'adj_fair_s{s}') for s in range(num_staff)]
                    
                    for s, staff in enumerate(staff_list):
                        model.Add(fair_shifts[s] == sum(shifts[(s, d)] for d in special_day_indices))
                        adj = fairness_adjustments.get(staff.name, 0) if fairness_adjustments else 0
                        model.Add(adj_fair[s] == fair_shifts[s] - adj)
                    
                    min_fair = model.NewIntVar(-num_days, num_days, 'min_fair')
                    max_fair = model.NewIntVar(0, num_days, 'max_fair')
                    model.AddMinEquality(min_fair, adj_fair)
                    model.AddMaxEquality(max_fair, adj_fair)
                    fair_diff = model.NewIntVar(0, num_days, 'fair_diff')
                    model.Add(fair_diff == max_fair - min_fair)

                    c2 = model.Add(fair_diff <= fairness_tolerance)
                    self.constraint_tags[c2.Index()] = f"特別日回数の公平性 (許容差: {fairness_tolerance}回)"

        model.Minimize(total_penalty)
        
    def _add_dispersion_penalty(self, model, shifts, staff_list, day_list, fairness_group, past_schedules):
        num_staff = len(staff_list)
        categories = {cat for cat in fairness_group}
        
        initial_penalties = defaultdict(lambda: defaultdict(int))
        today = day_list[0]['date']
        
        if past_schedules:
            for date_str, staff_names in past_schedules.items():
                past_date = datetime.date.fromisoformat(date_str)
                if (today - past_date).days > 90:
                    continue
                
                day_info = {
                    'weekday': weekdays_jp[past_date.weekday()],
                    'is_national_holiday': past_date in self.jp_holidays
                }
                
                for staff_name in staff_names:
                    staff_obj = self.staff_manager.get_staff_by_name(staff_name)
                    if staff_obj and staff_obj in staff_list:
                        for cat in self._get_date_categories(day_info, categories):
                            initial_penalties[staff_name][cat] += 1
        
        for staff_name in initial_penalties:
            for cat in initial_penalties[staff_name]:
                initial_penalties[staff_name][cat] *= 30

        penalty_vars = {}
        for s, staff in enumerate(staff_list):
            penalty_vars[s] = {}
            for cat in categories:
                initial_p = initial_penalties[staff.name][cat]
                penalty_vars[s][cat] = model.NewIntVar(0, 10000, f'penalty_s{s}_{cat}_d_start')
                model.Add(penalty_vars[s][cat] == initial_p)

        total_dispersion_penalty = model.NewIntVar(0, 1000000, 'dispersion_penalty')
        all_day_penalties = []

        for d, day_info in enumerate(day_list):
            day_categories = self._get_date_categories(day_info, categories)
            
            for s in range(num_staff):
                for cat in day_categories:
                    term = model.NewIntVar(0, 10000, f'p_term_s{s}_d{d}_{cat}')
                    model.Add(term == penalty_vars[s][cat]).OnlyEnforceIf(shifts[(s,d)])
                    model.Add(term == 0).OnlyEnforceIf(shifts[(s,d)].Not())
                    all_day_penalties.append(term)

            if d < len(day_list) - 1:
                for s in range(num_staff):
                    for cat in categories:
                        next_day_penalty_expr = penalty_vars[s][cat] - 1
                        if cat in day_categories:
                            next_day_penalty_expr += shifts[(s, d)] * 60
                        
                        next_day_penalty_var = model.NewIntVar(-10000, 10000, f'temp_penalty_s{s}_{cat}_d{d+1}')
                        model.Add(next_day_penalty_var == next_day_penalty_expr)
                        
                        non_negative_penalty = model.NewIntVar(0, 10000, f'penalty_s{s}_{cat}_d{d+1}')
                        model.AddMaxEquality(non_negative_penalty, [next_day_penalty_var, 0])
                        
                        penalty_vars[s][cat] = non_negative_penalty

        model.Add(total_dispersion_penalty == sum(all_day_penalties))
        return total_dispersion_penalty

    def _get_date_categories(self, day_info, target_categories):
        cats = set()
        weekday = day_info['weekday']
        if weekday in target_categories:
            cats.add(weekday)
        if day_info.get('is_national_holiday', False) and '祝' in target_categories:
            cats.add('祝')
        return cats

    def _add_solution_prohibition_constraint(self, model, shifts, solution):
        terms = []
        for s in range(len(self.all_staff)):
            for d in range(len(self.calendar_data)):
                if solution['raw_shifts'][(s, d)] == 1:
                    terms.append(shifts[(s, d)].Not())
                else:
                    terms.append(shifts[(s, d)])
        model.AddBoolOr(terms)

    def _create_solution_from_solver(self, solver, shifts, staff_list, day_list, fairness_group):
        schedule = {}
        raw_shifts_map = {}
        staff_map = {s.name: s for s in staff_list}

        for d, day_info in enumerate(day_list):
            date_obj = day_info['date']
            schedule[date_obj] = []
            for s, staff in enumerate(staff_list):
                is_working = solver.Value(shifts[(s, d)])
                raw_shifts_map[(s, d)] = is_working
                if is_working:
                    schedule[date_obj].append(staff_map[staff.name])

        counts = {staff.name: sum(solver.Value(shifts[s, d]) for d in range(len(day_list))) for s, staff in enumerate(staff_list)}
        fairness_counts = self._calculate_fairness_group_counts(schedule, fairness_group)

        return {
            "schedule": schedule,
            "counts": counts,
            "fairness_group_counts": fairness_counts,
            "raw_shifts": raw_shifts_map,
        }
        
    def _analyze_infeasibility(self, solver):
        assumptions = solver.SufficientAssumptionsForInfeasibility()
        if not assumptions:
            return "解が見つかりませんでした。原因の特定も困難です。ルールを緩めてみてください。"

        report_lines = ["シフトが見つかりませんでした。以下のルールが衝突している可能性があります："]
        for assumption_index in assumptions:
            tag = self.constraint_tags.get(assumption_index)
            if tag:
                report_lines.append(f"・ {tag}")
            else:
                lit = cp_model.Literal(assumption_index)
                tag = self.constraint_tags.get(lit.Var(), f"不明なルール({lit.Name()})")
                report_lines.append(f"・ {tag}")

        return "\n".join(report_lines)

    def _get_shift_range_for_day(self, day_info: dict, config) -> tuple[int, int]:
        if isinstance(config, int):
            return (config, config)
        if 'min' in config and 'max' in config and isinstance(config['min'], int):
            return (config['min'], config['max'])
        day_key = '祝' if day_info.get('is_national_holiday', False) and '祝' in config else day_info['weekday']
        day_setting = config.get(day_key, {'min': 1, 'max': 1})
        return (day_setting.get('min', 1), day_setting.get('max', 1))

    def _calculate_fairness_group_counts(self, schedule: dict, fairness_group: set) -> dict:
        if fairness_group is None: fairness_group = set()
        fairness_counts = {s.name: 0 for s in self.all_staff}
        for date_obj, staff_list in schedule.items():
            day_info = next((d for d in self.calendar_data if d['date'] == date_obj), None)
            if not day_info: continue
            is_fairness_day = ('祝' in fairness_group and day_info.get('is_national_holiday', False)) or \
                              (day_info['weekday'] in fairness_group)
            if is_fairness_day:
                for staff in staff_list:
                    if staff.name in fairness_counts:
                        fairness_counts[staff.name] += 1
        return fairness_counts

    def _generate_fixed_shifts_from_rules(self, rules: list, year: int, month: int) -> dict:
        rule_based_assignments = {}
        if not rules: return rule_based_assignments
        last_day_of_month = calendar.monthrange(year, month)[1]
        last_week_dates = {}
        for i in range(7):
            day = last_day_of_month - i
            if day > 0:
                date_obj = datetime.date(year, month, day)
                weekday_idx = date_obj.weekday()
                if weekday_idx not in last_week_dates:
                    last_week_dates[weekday_idx] = date_obj
            if len(last_week_dates) == 7: break
        weekday_counters = [0] * 7
        for day_info in self.calendar_data:
            current_date = day_info['date']
            if self.ignore_rules_on_holidays and current_date in self.jp_holidays: continue
            weekday_idx = current_date.weekday()
            weekday_counters[weekday_idx] += 1
            current_week_number = weekday_counters[weekday_idx]
            for rule in rules:
                is_last_week_match = (rule.week_number == 5 and rule.weekday_index == weekday_idx and last_week_dates.get(weekday_idx) == current_date)
                if (rule.week_number == current_week_number and rule.weekday_index == weekday_idx) or is_last_week_match:
                    if current_date in rule_based_assignments:
                        rule_based_assignments[current_date].append(rule.staff)
                    else:
                        rule_based_assignments[current_date] = [rule.staff]
        return rule_based_assignments

    def _generate_vacations_from_rules(self, rules: list, year: int, month: int) -> dict:
        rule_based_vacations = {}
        if not rules: return rule_based_vacations
        last_day_of_month = calendar.monthrange(year, month)[1]
        last_week_dates = {}
        for i in range(7):
            day = last_day_of_month - i
            if day > 0:
                date_obj = datetime.date(year, month, day)
                weekday_idx = date_obj.weekday()
                if weekday_idx not in last_week_dates:
                    last_week_dates[weekday_idx] = date_obj
            if len(last_week_dates) == 7: break
        weekday_counters = [0] * 7
        for day_info in self.calendar_data:
            current_date = day_info['date']
            if self.ignore_rules_on_holidays and current_date in self.jp_holidays: continue
            weekday_idx = current_date.weekday()
            weekday_counters[weekday_idx] += 1
            current_week_number = weekday_counters[weekday_idx]
            for rule in rules:
                is_last_week_match = (rule.week_number == 5 and rule.weekday_index == weekday_idx and last_week_dates.get(weekday_idx) == current_date)
                if (rule.week_number == current_week_number and rule.weekday_index == weekday_idx) or is_last_week_match:
                    staff_name = rule.staff_name
                    if staff_name not in rule_based_vacations:
                        rule_based_vacations[staff_name] = set()
                    rule_based_vacations[staff_name].add(current_date)
        return rule_based_vacations
    
    # ★★★ 削除: _generate_all_staff_impossible_dates メソッド ★★★

def generate_calendar_with_holidays(year: int, month: int):
    month_calendar = []
    jp_holidays = holidays.JP(years=year)
    _, num_days = calendar.monthrange(year, month)
    for day_num in range(1, num_days + 1):
        current_date = datetime.date(year, month, day_num)
        weekday_index = current_date.weekday()
        weekday_str = weekdays_jp[weekday_index]
        is_national_holiday = current_date in jp_holidays
        is_weekend = weekday_index >= 5
        is_holiday = is_national_holiday or is_weekend
        month_calendar.append({
            'date': current_date, 'weekday': weekday_str,
            'is_holiday': is_holiday, 'is_national_holiday': is_national_holiday,})
    return month_calendar
    
class SettingsManager:
    def __init__(self, history_dir: str = "shift_history"):
        self.staff_manager = StaffManager()
        self.rule_based_fixed_shifts: list[RuleBasedFixedShift] = []
        self.rule_based_vacations: list[RuleBasedVacation] = []
        self.min_interval = 2
        self.max_consecutive_days = 5
        self.shifts_per_day: dict = {'min': 1, 'max': 1}
        self.max_solutions = 1
        self.fairness_tolerance = 1 
        self.fairness_group: set[str] = {"土", "日", "祝"}
        self.ignore_rules_on_holidays: bool = False
        self.avoid_consecutive_same_weekday: bool = False
        self.disperse_duties: bool = True
        self.excel_title: str = "日直・当直予定表"
        self.history_dir = history_dir

        if not os.path.exists(self.history_dir):
            try:
                os.makedirs(self.history_dir)
                print(f"履歴保存用ディレクトリを初期化時に作成しました: {self.history_dir}")
            except Exception as e:
                print(f"警告: 履歴保存用ディレクトリの作成に失敗しました: {e}")

    def to_dict(self) -> dict:
        staff_list_dict = []
        for staff in self.staff_manager.get_all_staff():
            # ★★★ 修正: rule_based_impossible_days を削除 ★★★
            staff_list_dict.append({
                "name": staff.name, "color_code": staff.color_code,
                "is_active": staff.is_active,
                "impossible_weekdays": list(staff.impossible_weekdays)})
        rules_fixed_dict = [{"week": r.week_number, "weekday": r.weekday_index, "staff_name": r.staff.name} for r in self.rule_based_fixed_shifts]
        rules_vacation_dict = [{"week": r.week_number, "weekday": r.weekday_index, "staff_name": r.staff_name} for r in self.rule_based_vacations]
        
        general_settings_dict = {
            "min_interval": self.min_interval, 
            "max_consecutive_days": self.max_consecutive_days,
            "shifts_per_day": self.shifts_per_day,
            "max_solutions": self.max_solutions, 
            "fairness_tolerance": self.fairness_tolerance,
            "fairness_group": list(self.fairness_group),
            "ignore_rules_on_holidays": self.ignore_rules_on_holidays,
            "avoid_consecutive_same_weekday": self.avoid_consecutive_same_weekday,
            "disperse_duties": self.disperse_duties,
            "excel_title": self.excel_title
        }
            
        return {"staff": staff_list_dict, "rule_based_fixed_shifts": rules_fixed_dict,
                "rule_based_vacations": rules_vacation_dict, "general_settings": general_settings_dict}
                
    @classmethod
    def from_dict(cls, data: dict) -> 'SettingsManager':
        settings = cls()
        staff_data = data.get("staff", [])
        for s_data in staff_data:
            # ★★★ 修正: rule_based_impossible_days を削除 ★★★
            staff = Staff(name=s_data["name"], 
                          color_code=s_data["color_code"],
                          is_active=s_data.get("is_active", True),
                          impossible_weekdays=set(s_data.get("impossible_weekdays", [])))
            settings.staff_manager.add_or_update_staff(staff)
        rules_fixed_data = data.get("rule_based_fixed_shifts", [])
        for r_data in rules_fixed_data:
            staff_obj = settings.staff_manager.get_staff_by_name(r_data["staff_name"])
            if staff_obj:
                rule = RuleBasedFixedShift(week_number=r_data["week"], weekday_index=r_data["weekday"], staff=staff_obj)
                settings.rule_based_fixed_shifts.append(rule)
        rules_vacation_data = data.get("rule_based_vacations", [])
        for r_data in rules_vacation_data:
            rule = RuleBasedVacation(week_number=r_data["week"], weekday_index=r_data["weekday"], staff_name=r_data["staff_name"])
            settings.rule_based_vacations.append(rule)
            
        general_settings_data = data.get("general_settings", {})
        settings.min_interval = general_settings_data.get("min_interval", 2)
        settings.max_consecutive_days = general_settings_data.get("max_consecutive_days", 5)
        settings.shifts_per_day = general_settings_data.get("shifts_per_day", {'min': 1, 'max': 1})
        settings.max_solutions = general_settings_data.get("max_solutions", 1)
        settings.fairness_tolerance = general_settings_data.get("fairness_tolerance", 1)
        settings.fairness_group = set(general_settings_data.get("fairness_group", {"土", "日", "祝"}))
        settings.ignore_rules_on_holidays = general_settings_data.get("ignore_rules_on_holidays", False)
        settings.avoid_consecutive_same_weekday = general_settings_data.get("avoid_consecutive_same_weekday", False)
        settings.disperse_duties = general_settings_data.get("disperse_duties", True)
        settings.excel_title = general_settings_data.get("excel_title", "日直・当直予定表")
        
        return settings
    
    def save_to_json(self, filepath: str):
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(self.to_dict(), f, indent=4, ensure_ascii=False)
            print(f"設定を '{filepath}' に保存しました。")
        except Exception as e:
            print(f"設定の保存中にエラーが発生しました: {e}")
    @classmethod
    def load_from_json(self, filepath: str) -> bool:
        """
        JSONファイルから設定を読み込み、現在のインスタンスに適用する。
        成功した場合はTrue、失敗した場合はFalseを返す。
        """
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # from_dictを使って、読み込んだデータから新しいインスタンスを作成
            new_settings = SettingsManager.from_dict(data)
            
            # 新しいインスタンスのプロパティを現在のインスタンスにコピー
            # self.history_dir は上書きしない！
            self.staff_manager = new_settings.staff_manager
            self.rule_based_fixed_shifts = new_settings.rule_based_fixed_shifts
            self.rule_based_vacations = new_settings.rule_based_vacations
            self.min_interval = new_settings.min_interval
            self.max_consecutive_days = new_settings.max_consecutive_days
            self.shifts_per_day = new_settings.shifts_per_day
            self.max_solutions = new_settings.max_solutions
            self.fairness_tolerance = new_settings.fairness_tolerance
            self.fairness_group = new_settings.fairness_group
            self.ignore_rules_on_holidays = new_settings.ignore_rules_on_holidays
            self.avoid_consecutive_same_weekday = new_settings.avoid_consecutive_same_weekday
            self.disperse_duties = new_settings.disperse_duties
            self.excel_title = new_settings.excel_title

            print(f"'{filepath}' から設定を読み込みました。")
            return True
        except FileNotFoundError:
            print(f"設定ファイル '{filepath}' が見つかりません。")
            return False
        except Exception as e:
            print(f"設定の読み込み中にエラーが発生しました: {e}")
            return False

    def history_exists(self, year: int, month: int) -> bool:
        """指定された年月の履歴ファイルが存在するかどうかをチェックする"""
        filepath = os.path.join(self.history_dir, f"history_{year}-{month:02d}.json")
        return os.path.exists(filepath)
    
    def save_history(self, year: int, month: int, solution: dict):
        filepath = os.path.join(self.history_dir, f"history_{year}-{month:02d}.json")
        schedule_for_json = []
        for date_obj, staff_list in solution["schedule"].items():
            schedule_for_json.append({"date": date_obj.isoformat(), "staff_names": [s.name for s in staff_list]})
        history_data = {
            "year": year, "month": month, "schedule": schedule_for_json,
            "counts": solution["counts"],
            "fairness_group_counts": solution["fairness_group_counts"]}
        try:
            # 1. 保存先ディレクトリのパスを取得
            history_directory = os.path.dirname(filepath)
            
            # 2. ディレクトリが存在しない場合、作成する
            #    os.makedirs(..., exist_ok=True) は、途中のフォルダも含めて再帰的に作成し、
            #    既に存在していてもエラーにならない便利なオプションです。
            if not os.path.exists(history_directory):
                os.makedirs(history_directory, exist_ok=True)
                print(f"履歴保存用ディレクトリを作成しました: {history_directory}")
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(history_data, f, indent=4, ensure_ascii=False)
            print(f"履歴を '{filepath}' に保存しました。")
            return True
        except Exception as e:
            print(f"履歴の保存中にエラーが発生しました: {e}")
            return False
    def load_history(self, year: int, month: int) -> dict | None:
        filepath = os.path.join(self.history_dir, f"history_{year}-{month:02d}.json")
        if not os.path.exists(filepath): return None
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                history_data = json.load(f)
            print(f"履歴を読み込みました。")
            return history_data
        except Exception as e:
            print(f"履歴の読み込み中にエラーが発生しました: {e}")
            return None

import sys
from functools import partial
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QFormLayout,
    QTableWidget, QTableWidgetItem, QLineEdit, QPushButton, QCheckBox,
    QHeaderView, QGroupBox, QGridLayout, QComboBox, QListWidget,
    QMessageBox, QColorDialog
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor

from core_engine import weekdays_jp, Staff, SettingsManager

class StaffConfigTab(QWidget):
    def __init__(self, settings_manager: SettingsManager):
        super().__init__()
        self.settings_manager = settings_manager
        self._init_ui()
        self._connect_signals()

    def set_settings_manager(self, settings_manager: SettingsManager):
        self.settings_manager = settings_manager
        self.load_staff_list()

    def _connect_signals(self):
        self.table.itemSelectionChanged.connect(self._on_staff_selected)
        self.add_button.clicked.connect(self._add_or_update_staff)
        self.delete_button.clicked.connect(self._delete_staff)
        self.clear_form_button.clicked.connect(self._clear_form)
        self.color_picker_button.clicked.connect(self._open_color_picker)

    def _init_ui(self):
        main_layout = QVBoxLayout(self)

        self.table = QTableWidget()
        # ★★★ 修正: カラム数を4に減らし、ヘッダーラベルを変更 ★★★
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["稼働中", "名前", "カラー", "固定の不可曜日"])
        
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)

        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        main_layout.addWidget(self.table)

        form_layout = QFormLayout()
        self.name_input = QLineEdit()

        self.color_input = QLineEdit()
        self.color_input.setPlaceholderText("#RRGGBB形式 (例: #FFADAD)")
        self.color_picker_button = QPushButton("色を選択...")

        color_layout = QHBoxLayout()
        color_layout.setContentsMargins(0, 0, 0, 0)
        
        color_layout.addWidget(self.color_picker_button)
        color_layout.addWidget(self.color_input)

        self.weekdays_layout = QHBoxLayout()
        self.weekday_checkboxes = {}
        for day in weekdays_jp:
            checkbox = QCheckBox(day)
            self.weekday_checkboxes[day] = checkbox
            self.weekdays_layout.addWidget(checkbox)

        form_layout.addRow("名前:", self.name_input)
        form_layout.addRow("カラー:", color_layout)
        form_layout.addRow("固定の不可曜日:", self.weekdays_layout)
        main_layout.addLayout(form_layout)

        # ★★★ 削除: rule_groupbox 全体 ★★★

        button_layout = QHBoxLayout()
        self.add_button = QPushButton("スタッフを追加/更新")
        self.delete_button = QPushButton("選択したスタッフを削除")
        self.clear_form_button = QPushButton("フォームをクリア")
        button_layout.addStretch()
        button_layout.addWidget(self.clear_form_button)
        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.delete_button)
        main_layout.addLayout(button_layout)
    
    def _open_color_picker(self):
        current_color_str = self.color_input.text()
        initial_color = QColor(current_color_str) if QColor.isValidColor(current_color_str) else Qt.GlobalColor.white
        color = QColorDialog.getColor(initial_color, self, "色の選択")
        if color.isValid():
            self.color_input.setText(color.name())

    def load_staff_list(self):
        self.table.blockSignals(True)
        self.table.setRowCount(0)
        staff_list = self.settings_manager.staff_manager.get_all_staff()
        
        for staff in sorted(staff_list, key=lambda s: s.name):
            row_position = self.table.rowCount()
            self.table.insertRow(row_position)
            
            active_checkbox = QCheckBox()
            active_checkbox.setChecked(staff.is_active)
            active_checkbox.stateChanged.connect(partial(self._toggle_staff_active, staff.name))
            
            wrapper = QWidget()
            layout = QHBoxLayout(wrapper)
            layout.addWidget(active_checkbox)
            layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
            layout.setContentsMargins(0,0,0,0)
            self.table.setCellWidget(row_position, 0, wrapper)

            self.table.setItem(row_position, 1, QTableWidgetItem(staff.name))
            
            color_item = QTableWidgetItem(staff.color_code)
            try:
                color_item.setBackground(QColor(staff.color_code))
            except Exception: pass
            self.table.setItem(row_position, 2, color_item)
            
            impossible_days_str = ", ".join(sorted(list(staff.impossible_weekdays)))
            self.table.setItem(row_position, 3, QTableWidgetItem(impossible_days_str))

            # ★★★ 削除: rule_days_str の処理 ★★★
        
        self.table.blockSignals(False)

    def _toggle_staff_active(self, staff_name, state):
        staff = self.settings_manager.staff_manager.get_staff_by_name(staff_name)
        if staff:
            staff.is_active = (state == Qt.CheckState.Checked.value)
            print(f"スタッフ '{staff.name}' の稼働状態が '{staff.is_active}' に変更されました。")

    def _on_staff_selected(self):
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            return

        row = selected_rows[0].row()
        staff_name_item = self.table.item(row, 1) 
        if not staff_name_item: return
        
        staff_name = staff_name_item.text()
        staff = self.settings_manager.staff_manager.get_staff_by_name(staff_name)

        if not staff:
            return

        self._clear_form(clear_selection=False)
        self.name_input.setText(staff.name)
        self.name_input.setReadOnly(True)
        self.color_input.setText(staff.color_code)

        for day, checkbox in self.weekday_checkboxes.items():
            checkbox.setChecked(day in staff.impossible_weekdays)
        
        # ★★★ 削除: rule_list_widget のクリア処理 ★★★

    def _add_or_update_staff(self):
        name = self.name_input.text().strip()
        if not name:
            QMessageBox.warning(self, "入力エラー", "スタッフ名を入力してください。")
            return
        
        color = self.color_input.text().strip()
        if not QColor.isValidColor(color) or not color.startswith('#'):
            QMessageBox.warning(self, "入力エラー", "カラーコードが不正な形式です。'#RRGGBB' の形式で入力してください。 (例: #ffadad)")
            return

        impossible_weekdays = {day for day, cb in self.weekday_checkboxes.items() if cb.isChecked()}
        
        # ★★★ 削除: rule_based_impossible_days の取得処理 ★★★
        
        existing_staff = self.settings_manager.staff_manager.get_staff_by_name(name)
        is_active = existing_staff.is_active if existing_staff else True

        # ★★★ 修正: Staffのコンストラクタから rule_based_impossible_days を削除 ★★★
        new_staff = Staff(
            name=name,
            color_code=color.lower(),
            is_active=is_active,
            impossible_weekdays=impossible_weekdays
        )

        self.settings_manager.staff_manager.add_or_update_staff(new_staff)
        self.load_staff_list()
        self._clear_form()

    def _delete_staff(self):
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, "選択エラー", "削除するスタッフをリストから選択してください。")
            return

        row = selected_rows[0].row()
        staff_name_item = self.table.item(row, 1)
        if not staff_name_item: return

        staff_name = staff_name_item.text()

        reply = QMessageBox.question(self, '削除の確認', 
                                     f"本当にスタッフ '{staff_name}' を削除しますか？",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, 
                                     QMessageBox.StandardButton.No)

        if reply == QMessageBox.StandardButton.Yes:
            self.settings_manager.staff_manager.remove_staff_by_name(staff_name)
            self.load_staff_list()
            self._clear_form()

    def _clear_form(self, clear_selection=True):
        self.name_input.clear()
        self.name_input.setReadOnly(False)
        self.color_input.clear()
        for checkbox in self.weekday_checkboxes.values():
            checkbox.setChecked(False)
        # ★★★ 削除: rule_list_widget のクリア処理 ★★★
        if clear_selection:
            self.table.clearSelection()

    # ★★★ 削除: _add_impossible_rule, _delete_impossible_rule, _format_rule, _parse_rule メソッド ★★★