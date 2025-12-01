import calendar
import datetime
import json
import os
from dataclasses import dataclass
from typing import List, Dict, Set, Tuple, Optional
from collections import defaultdict

import holidays
from ortools.sat.python import cp_model


weekdays_jp = ("月", "火", "水", "木", "金", "土", "日")


class Staff:
    def __init__(self, name: str, color_code: str, impossible_weekdays: Set[str] | None = None, is_active: bool = True):
        if not isinstance(name, str) or not name:
            raise ValueError("invalid staff name")
        if not isinstance(color_code, str) or not color_code.startswith('#'):
            raise ValueError("invalid color code (expect startswith '#')")
        self.name = name
        self.color_code = color_code
        self.impossible_weekdays = set(impossible_weekdays) if impossible_weekdays else set()
        self.is_active = is_active

    def is_available(self, weekday: str) -> bool:
        return weekday not in self.impossible_weekdays

    def __repr__(self) -> str:
        return (f"Staff(name={self.name}, color={self.color_code}, active={self.is_active}, "
                f"impossible={sorted(list(self.impossible_weekdays))})")


class StaffManager:
    def __init__(self):
        self.staff_map: Dict[str, Staff] = {}

    def add_or_update_staff(self, staff: Staff):
        self.staff_map[staff.name] = staff

    def remove_staff_by_name(self, name: str) -> bool:
        if name in self.staff_map:
            del self.staff_map[name]
            return True
        return False

    def get_all_staff(self) -> List[Staff]:
        return list(self.staff_map.values())

    def get_active_staff(self) -> List[Staff]:
        return [s for s in self.staff_map.values() if s.is_active]

    def __len__(self):
        return len(self.staff_map)

    def get_staff_by_name(self, name: str) -> Optional[Staff]:
        return self.staff_map.get(name)


@dataclass(frozen=True)
class RuleBasedFixedShift:
    week_number: int  # 1..5 (5=last)
    weekday_index: int  # 0..6 (Mon..Sun)
    staff: Staff


@dataclass(frozen=True)
class RuleBasedVacation:
    week_number: int
    weekday_index: int
    staff_name: str


def generate_calendar_with_holidays(year: int, month: int) -> List[dict]:
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
            'date': current_date,
            'weekday': weekday_str,
            'is_holiday': is_holiday,
            'is_national_holiday': is_national_holiday,
        })
    return month_calendar


class ShiftScheduler:
    def __init__(self, staff_manager: StaffManager, calendar_data: List[dict], ignore_rules_on_holidays: bool = False):
        self.staff_manager = staff_manager
        self.calendar_data = calendar_data
        self.all_staff = self.staff_manager.get_active_staff()
        self.ignore_rules_on_holidays = ignore_rules_on_holidays
        self.jp_holidays = holidays.JP(years=self.calendar_data[0]['date'].year)
        self.constraint_tags: Dict[int, str] = {}

    def solve(self,
               shifts_per_day: int | dict = 1,
              min_interval: int = 2,
              max_consecutive_days: int = 5,
              max_solutions: int = 1,
              last_month_end_dates: dict | None = None,
              prev_month_consecutive_days: dict | None = None,
              last_week_assignments: dict | None = None,
              avoid_consecutive_same_weekday: bool = False,
              no_shift_dates: List[datetime.date] | None = None,
              manual_fixed_shifts: dict | None = None,
              rule_based_fixed_shifts: List[RuleBasedFixedShift] | None = None,
              vacations: dict | None = None,
              rule_based_vacations: List[RuleBasedVacation] | None = None,
              fairness_group: set | None = None,
              total_adjustments: dict | None = None,
              fairness_adjustments: dict | None = None,
              fairness_tolerance: int = 1,
              disperse_duties: bool = True,
               past_schedules: dict | None = None,
               fairness_as_hard: bool = True,
               fallback_soft_on_infeasible: bool = True,
               **kwargs
               ) -> List[dict] | str:

        staff_list = self.all_staff
        if not staff_list:
            return []

        self.past_schedules = past_schedules or {}
        self._fairness_as_hard = fairness_as_hard
        # Collect solutions to allow duplicate-prevention and final return
        found_solutions: List[dict] = []
        for _ in range(max_solutions):
            model = cp_model.CpModel()
            self.constraint_tags = {}
            shifts = self._define_variables(model, staff_list, self.calendar_data)

            self._add_hard_constraints(model, shifts, staff_list, self.calendar_data,
                                       no_shift_dates, shifts_per_day,
                                       rule_based_vacations, vacations,
                                       min_interval, max_consecutive_days,
                                       last_month_end_dates, prev_month_consecutive_days,
                                       fairness_group, avoid_consecutive_same_weekday,
                                       last_week_assignments,
                                       manual_fixed_shifts,
                                       rule_based_fixed_shifts)

            _, fixed_penalty = self._add_soft_constraints(model, shifts, staff_list, self.calendar_data,
                                                          rule_based_fixed_shifts, None)
            dispersion_penalty = 0
            if disperse_duties and fairness_group:
                # ORIGINE 相当: 直近のスケジュール傾向を考慮し、特定カテゴリの偏りを抑える
                dispersion_penalty = self._add_dispersion_penalty(
                    model, shifts, staff_list, self.calendar_data, fairness_group, self.past_schedules
                )
            self._add_fairness_objective(
                model, shifts, staff_list, self.calendar_data,
                total_adjustments or {}, fairness_adjustments or {},
                fairness_tolerance, fairness_group or set(),
                fixed_penalty, dispersion_penalty
            )

            # Prohibit already found solutions before solving
            for sol in found_solutions:
                self._add_solution_prohibition_constraint(model, shifts, sol)

            solver = cp_model.CpSolver()
            status = solver.Solve(model)

            if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
                solution = self._create_solution_from_solver(solver, shifts, staff_list, self.calendar_data, fairness_group or set())
                found_solutions.append(solution)
            else:
                if status == cp_model.INFEASIBLE:
                    # 公平性がハードかつフォールバック許可時は、ソフトにして再実行
                    if fallback_soft_on_infeasible and fairness_as_hard and (fairness_group or set()):
                        alt = self.solve(
                            shifts_per_day=shifts_per_day,
                            min_interval=min_interval,
                            max_consecutive_days=max_consecutive_days,
                            max_solutions=max_solutions,
                            last_month_end_dates=last_month_end_dates,
                            prev_month_consecutive_days=prev_month_consecutive_days,
                            last_week_assignments=last_week_assignments,
                            avoid_consecutive_same_weekday=avoid_consecutive_same_weekday,
                            no_shift_dates=no_shift_dates,
                            manual_fixed_shifts=manual_fixed_shifts,
                            rule_based_fixed_shifts=rule_based_fixed_shifts,
                            vacations=vacations,
                            rule_based_vacations=rule_based_vacations,
                            fairness_group=fairness_group,
                            total_adjustments=total_adjustments,
                            fairness_adjustments=fairness_adjustments,
                            fairness_tolerance=fairness_tolerance,
                            disperse_duties=disperse_duties,
                            past_schedules=self.past_schedules,
                            fairness_as_hard=False,
                            fallback_soft_on_infeasible=False
                        )
                        if isinstance(alt, list) and alt:
                            for sdict in alt:
                                sdict['generation_note'] = '公平性（特別日）をソフトに緩和して生成'
                            return alt
                    try:
                        return self._analyze_infeasibility(solver)
                    except Exception:
                        return "シフトが見つかりませんでした（制約の衝突）"
                break

        return found_solutions

    def _define_variables(self, model, staff_list, day_list):
        shifts = {}
        for s in range(len(staff_list)):
            for d in range(len(day_list)):
                shifts[(s, d)] = model.NewBoolVar(f'shift_s{s}_d{d}')
        return shifts

    def _get_shift_range_for_day(self, day_info: dict, config) -> Tuple[int, int]:
        if isinstance(config, int):
            return (config, config)
        if 'min' in config and 'max' in config and isinstance(config.get('min'), int):
            return (config.get('min', 1), config.get('max', 1))
        # keyed by weekday or '逾・ for holiday
        key = '祝' if (day_info.get('is_national_holiday', False) and '祝' in config) else day_info['weekday']
        setting = config.get(key, {'min': 1, 'max': 1})
        return (setting.get('min', 1), setting.get('max', 1))

    def _generate_fixed_shifts_from_rules(self, rules: List[RuleBasedFixedShift] | None, year: int, month: int) -> Dict[datetime.date, List[Staff]]:
        result: Dict[datetime.date, List[Staff]] = {}
        if not rules:
            return result
        last_day = calendar.monthrange(year, month)[1]
        last_week_dates: Dict[int, datetime.date] = {}
        for i in range(7):
            d = last_day - i
            if d > 0:
                date_obj = datetime.date(year, month, d)
                wd = date_obj.weekday()
                if wd not in last_week_dates:
                    last_week_dates[wd] = date_obj
            if len(last_week_dates) == 7:
                break
        counters = [0] * 7
        for day in self.calendar_data:
            current_date = day['date']
            if self.ignore_rules_on_holidays and current_date in self.jp_holidays:
                continue
            wd = current_date.weekday()
            counters[wd] += 1
            current_week = counters[wd]
            for r in rules:
                is_last = (r.week_number == 5 and r.weekday_index == wd and last_week_dates.get(wd) == current_date)
                if (r.week_number == current_week and r.weekday_index == wd) or is_last:
                    result.setdefault(current_date, []).append(r.staff)
        return result

    def _generate_vacations_from_rules(self, rules: List[RuleBasedVacation] | None, year: int, month: int) -> Dict[str, Set[datetime.date]]:
        result: Dict[str, Set[datetime.date]] = {}
        if not rules:
            return result
        last_day = calendar.monthrange(year, month)[1]
        last_week_dates: Dict[int, datetime.date] = {}
        for i in range(7):
            d = last_day - i
            if d > 0:
                date_obj = datetime.date(year, month, d)
                wd = date_obj.weekday()
                if wd not in last_week_dates:
                    last_week_dates[wd] = date_obj
            if len(last_week_dates) == 7:
                break
        counters = [0] * 7
        for day in self.calendar_data:
            current_date = day['date']
            if self.ignore_rules_on_holidays and current_date in self.jp_holidays:
                continue
            wd = current_date.weekday()
            counters[wd] += 1
            current_week = counters[wd]
            for r in rules:
                is_last = (r.week_number == 5 and r.weekday_index == wd and last_week_dates.get(wd) == current_date)
                if (r.week_number == current_week and r.weekday_index == wd) or is_last:
                    result.setdefault(r.staff_name, set()).add(current_date)
        return result

    def _add_hard_constraints(self, model, shifts, staff_list, day_list,
                               no_shift_dates, shifts_per_day_config,
                               rule_based_vacations, vacations,
                               min_interval, max_consecutive_days,
                               last_month_end_dates, prev_month_consecutive_days,
                               fairness_group, avoid_consecutive_same_weekday,
                               last_week_assignments,
                               manual_fixed_shifts: dict | None = None,
                               rule_based_fixed_shifts_list: List[RuleBasedFixedShift] | None = None):

        num_staff = len(staff_list)
        year = day_list[0]['date'].year
        month = day_list[0]['date'].month
        first_day_of_month = day_list[0]['date']

        # Per-day capacity (with no-shift dates)
        for d, day_info in enumerate(day_list):
            date_obj = day_info['date']
            if no_shift_dates and date_obj in no_shift_dates:
                c = model.Add(sum(shifts[(s, d)] for s in range(num_staff)) == 0)
                try:
                    self.constraint_tags[c.Index()] = f"{date_obj.day}日の必要人数（不要日=0）"
                except Exception:
                    pass
                continue
            min_needed, max_needed = self._get_shift_range_for_day(day_info, shifts_per_day_config)
            c_need = model.AddLinearConstraint(sum(shifts[(s, d)] for s in range(num_staff)), min_needed, max_needed)
            try:
                self.constraint_tags[c_need.Index()] = f"{date_obj.day}日の必要人数（{min_needed}〜{max_needed}人）"
            except Exception:
                pass

        # Build vacation/fixed maps
        generated_vac = self._generate_vacations_from_rules(rule_based_vacations, year, month)
        manual_vacations = vacations or {}
        manual_fixed_shifts = manual_fixed_shifts or {}
        rule_fixed_from_rules = self._generate_fixed_shifts_from_rules(rule_based_fixed_shifts_list, year, month)

        manual_fixed_lookup = {(st.name, d) for d, lst in manual_fixed_shifts.items() for st in lst}
        rule_fixed_lookup = {(st.name, d) for d, lst in rule_fixed_from_rules.items() for st in lst}
        planned_fixed_lookup = manual_fixed_lookup | rule_fixed_lookup

        # Hard rules per staff per day (priority)
        for s, staff in enumerate(staff_list):
            for d, day_info in enumerate(day_list):
                date_obj = day_info['date']
                # Month-limited vacation (hard 0)
                if date_obj in manual_vacations.get(staff.name, set()):
                    c = model.Add(shifts[(s, d)] == 0)
                    try:
                        self.constraint_tags[c.Index()] = f"{staff.name}の{date_obj.day}日（月限定休暇）"
                    except Exception:
                        pass
                    continue
                # Month-limited fixed (hard 1)
                if (staff.name, date_obj) in manual_fixed_lookup:
                    c = model.Add(shifts[(s, d)] == 1)
                    try:
                        self.constraint_tags[c.Index()] = f"{staff.name}の{date_obj.day}日（月限定固定）"
                    except Exception:
                        pass
                    continue
                # Rule-based vacation (hard 0)
                if date_obj in generated_vac.get(staff.name, set()):
                    c = model.Add(shifts[(s, d)] == 0)
                    try:
                        self.constraint_tags[c.Index()] = f"{staff.name}の{date_obj.day}日（ルール休暇）"
                    except Exception:
                        pass
                    continue
                # Staff impossible weekday (weakest)
                is_holiday_and_ignored = self.ignore_rules_on_holidays and date_obj in self.jp_holidays
                if not staff.is_available(day_info['weekday']) and not is_holiday_and_ignored:
                    c = model.Add(shifts[(s, d)] == 0)
                    try:
                        self.constraint_tags[c.Index()] = f"{staff.name}の{day_info['weekday']}曜日の不可日"
                    except Exception:
                        pass

            # last-month carry over for min interval
            if last_month_end_dates and staff.name in last_month_end_dates:
                last_worked_date = last_month_end_dates[staff.name]
                days_since_last = (first_day_of_month - last_worked_date).days
                days_to_forbid = min_interval - days_since_last + 1
                for d in range(days_to_forbid):
                    if d < len(day_list):
                        date_obj2 = day_list[d]['date']
                        if (staff.name, date_obj2) in planned_fixed_lookup:
                            continue
                        c = model.Add(shifts[(s, d)] == 0)
                        try:
                            self.constraint_tags[c.Index()] = f"{staff.name}の{date_obj2.day}日の勤務不可（前月からの間隔）"
                        except Exception:
                            pass

            # Min interval (skip only when both ends are planned fixed)
            for d in range(len(day_list) - min_interval - 1):
                lit_d = shifts[(s, d)]
                lit_d1_not = shifts[(s, d + 1)].Not()
                for i in range(1, min_interval):
                    k = d + 1 + i
                    if k >= len(day_list):
                        continue
                    d_date = day_list[d]['date']
                    k_date = day_list[k]['date']
                    if (staff.name, d_date) in planned_fixed_lookup and (staff.name, k_date) in planned_fixed_lookup:
                        continue
                    c = model.AddImplication(shifts[(s, k)], lit_d.Not()).OnlyEnforceIf([lit_d, lit_d1_not])
                    try:
                        self.constraint_tags[c.Index()] = f"{staff.name}の{d_date.day}日からの休み間隔"
                    except Exception:
                        pass

            # Max consecutive (skip window only if it contains adjacent planned fixed pair)
            for d in range(len(day_list) - max_consecutive_days):
                window_indices = list(range(d, d + max_consecutive_days + 1))
                has_planned_pair = False
                for j in range(len(window_indices) - 1):
                    d1 = day_list[window_indices[j]]['date']
                    d2 = day_list[window_indices[j+1]]['date']
                    if (staff.name, d1) in planned_fixed_lookup and (staff.name, d2) in planned_fixed_lookup:
                        has_planned_pair = True
                        break
                if has_planned_pair:
                    continue
                window = [shifts[(s, i)] for i in window_indices]
                c = model.Add(sum(window) <= max_consecutive_days)
                try:
                    self.constraint_tags[c.Index()] = f"{staff.name}の{day_list[d]['date'].day}日からの最大連勤"
                except Exception:
                    pass

            if prev_month_consecutive_days and staff.name in prev_month_consecutive_days:
                consecutive = prev_month_consecutive_days[staff.name]
                if consecutive > 0:
                    remaining_days = max_consecutive_days - consecutive
                    if remaining_days < max_consecutive_days:
                        indices = list(range(remaining_days + 1))
                        has_planned_pair2 = False
                        for j in range(len(indices) - 1):
                            d1 = day_list[indices[j]]['date']
                            d2 = day_list[indices[j+1]]['date']
                            if (staff.name, d1) in planned_fixed_lookup and (staff.name, d2) in planned_fixed_lookup:
                                has_planned_pair2 = True
                                break
                        if not has_planned_pair2:
                            window = [shifts[(s, d)] for d in indices]
                            c = model.Add(sum(window) <= remaining_days)
                            try:
                                self.constraint_tags[c.Index()] = f"{staff.name}の月初の連勤制限"
                            except Exception:
                                pass

    def _add_soft_constraints(self, model, shifts, staff_list, day_list,
                              rule_based_fixed_shifts, manual_fixed_shifts):
        # Only rule-based fixed shifts are soft; month-limited fixed are hard
        year = day_list[0]['date'].year
        month = day_list[0]['date'].month
        rule_fixed = self._generate_fixed_shifts_from_rules(rule_based_fixed_shifts, year, month)

        penalty_literals: List[cp_model.IntVar] = []
        penalty_cost = model.NewIntVar(0, 0, 'empty_penalty')

        for date_obj, staff_obj_list in rule_fixed.items():
            try:
                d = [i for i, day in enumerate(day_list) if day['date'] == date_obj][0]
            except (ValueError, IndexError):
                continue
            for staff_obj in staff_obj_list:
                if staff_obj not in staff_list:
                    continue
                s = staff_list.index(staff_obj)
                lit = model.NewBoolVar(f"fixed_penalty_s{s}_d{d}")
                model.Add(shifts[(s, d)] == 0).OnlyEnforceIf(lit)
                model.Add(shifts[(s, d)] == 1).OnlyEnforceIf(lit.Not())
                penalty_literals.append(lit)

        if penalty_literals:
            penalty_cost = model.NewIntVar(0, len(penalty_literals), 'penalty_cost')
            model.Add(penalty_cost == sum(penalty_literals))

        return penalty_literals, penalty_cost
    def _add_fairness_objective(self, model, shifts, staff_list, day_list,
                                total_adjustments, fairness_adjustments,
                                fairness_tolerance, fairness_group,
                                fixed_shift_penalty, dispersion_penalty):
        num_staff = len(staff_list)
        num_days = len(day_list)
        fairness_penalty = model.NewIntVar(0, 1000000, 'fairness_penalty')
        model.Add(fairness_penalty == 0)
        total_penalty = model.NewIntVar(0, 1000000, 'total_penalty')
        model.Add(total_penalty == fixed_shift_penalty + dispersion_penalty + fairness_penalty)
        if num_staff > 1:
            total_shifts = [model.NewIntVar(0, num_days, f'total_s{s}') for s in range(num_staff)]
            adj_total = [model.NewIntVar(-num_days, num_days, f'adj_total_s{s}') for s in range(num_staff)]
            for s_idx, staff in enumerate(staff_list):
                model.Add(total_shifts[s_idx] == sum(shifts[(s_idx, d)] for d in range(num_days)))
                adj = total_adjustments.get(staff.name, 0) if total_adjustments else 0
                model.Add(adj_total[s_idx] == total_shifts[s_idx] - adj)
            min_total = model.NewIntVar(-num_days, num_days, 'min_total')
            max_total = model.NewIntVar(0, num_days, 'max_total')
            model.AddMinEquality(min_total, adj_total)
            model.AddMaxEquality(max_total, adj_total)
            total_diff = model.NewIntVar(0, num_days, 'total_diff')
            model.Add(total_diff == max_total - min_total)

            # 特定曜日/祝日の公平性（hard: 制約 / soft: ペナルティ）
            if fairness_group:
                special_day_indices = [
                    d for d, day_info in enumerate(day_list)
                    if ('祝' in fairness_group and day_info.get('is_national_holiday', False))
                    or (day_info.get('weekday') in fairness_group)
                ]

                if special_day_indices:
                    fair_shifts = [model.NewIntVar(0, len(special_day_indices), f'fair_s{s}') for s in range(num_staff)]
                    adj_fair = [model.NewIntVar(-num_days, num_days, f'adj_fair_s{s}') for s in range(num_staff)]

                    for s_idx, staff in enumerate(staff_list):
                        model.Add(fair_shifts[s_idx] == sum(shifts[(s_idx, d)] for d in special_day_indices))
                        adj = fairness_adjustments.get(staff.name, 0) if fairness_adjustments else 0
                        model.Add(adj_fair[s_idx] == fair_shifts[s_idx] - adj)

                    min_fair = model.NewIntVar(-num_days, num_days, 'min_fair')
                    max_fair = model.NewIntVar(0, num_days, 'max_fair')
                    model.AddMinEquality(min_fair, adj_fair)
                    model.AddMaxEquality(max_fair, adj_fair)
                    fair_diff = model.NewIntVar(0, num_days, 'fair_diff')
                    model.Add(fair_diff == max_fair - min_fair)

                    if self._fairness_as_hard:
                        c2 = model.Add(fair_diff <= fairness_tolerance)
                        try:
                            self.constraint_tags[c2.Index()] = f"特別日回数の公平性 (許容差: {fairness_tolerance}回)"
                        except Exception:
                            pass
                    else:
                        t = model.NewIntVar(-num_days, num_days, 'fair_over_tmp')
                        model.Add(t == fair_diff - fairness_tolerance)
                        over = model.NewIntVar(0, num_days, 'fair_over')
                        model.AddMaxEquality(over, [t, 0])
                        model.Add(fairness_penalty == over)

            # 目的関数: 総回数差 + 既存ペナルティ（固定/分散）を最小化
            model.Minimize(total_diff + total_penalty)
        else:
            model.Minimize(total_penalty)

    def _add_solution_prohibition_constraint(self, model, shifts, solution):
        terms = []
        num_staff = len(self.all_staff)
        num_days = len(self.calendar_data)
        for s in range(num_staff):
            for d in range(num_days):
                if solution['raw_shifts'][(s, d)] == 1:
                    terms.append(shifts[(s, d)].Not())
                else:
                    terms.append(shifts[(s, d)])
        model.AddBoolOr(terms)
    def _get_date_categories(self, day_info, target_categories):
        cats = set()
        weekday = day_info['weekday']
        if weekday in target_categories:
            cats.add(weekday)
        if day_info.get('is_national_holiday', False) and '祝' in target_categories:
            cats.add('祝')
        return cats
    def _add_dispersion_penalty(self, model, shifts, staff_list, day_list, fairness_group, past_schedules):
        num_staff = len(staff_list)
        categories = {cat for cat in fairness_group}

        # 初期ペナルティ: 過去90日以内の同カテゴリ実績を強めに重み付け
        initial_penalties = defaultdict(lambda: defaultdict(int))
        today = day_list[0]['date']

        if past_schedules:
            for date_str, staff_names in past_schedules.items():
                try:
                    past_date = datetime.date.fromisoformat(date_str)
                except Exception:
                    continue
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

        # スケール調整（累積に係数を掛ける）
        for staff_name in initial_penalties:
            for cat in initial_penalties[staff_name]:
                initial_penalties[staff_name][cat] *= 30

        # 進行とともに更新されるカテゴリ別ペナルティ変数
        penalty_vars = {}
        for s, staff in enumerate(staff_list):
            penalty_vars[s] = {}
            for cat in categories:
                initial_p = initial_penalties[staff.name][cat]
                v = model.NewIntVar(0, 10000, f'penalty_s{s}_{cat}_d_start')
                model.Add(v == initial_p)
                penalty_vars[s][cat] = v

        total_dispersion_penalty = model.NewIntVar(0, 1000000, 'dispersion_penalty')
        all_day_penalties: List[cp_model.IntVar] = []

        for d, day_info in enumerate(day_list):
            day_categories = self._get_date_categories(day_info, categories)
            for s in range(num_staff):
                for cat in day_categories:
                    term = model.NewIntVar(0, 10000, f'p_term_s{s}_d{d}_{cat}')
                    model.Add(term == penalty_vars[s][cat]).OnlyEnforceIf(shifts[(s, d)])
                    model.Add(term == 0).OnlyEnforceIf(shifts[(s, d)].Not())
                    all_day_penalties.append(term)

            # 翌日に向けてペナルティを減衰/更新
            if d < len(day_list) - 1:
                for s in range(num_staff):
                    for cat in categories:
                        next_expr = penalty_vars[s][cat] - 1
                        if cat in day_categories:
                            next_expr += shifts[(s, d)] * 60
                        temp = model.NewIntVar(-10000, 10000, f'temp_penalty_s{s}_{cat}_d{d+1}')
                        model.Add(temp == next_expr)
                        nonneg = model.NewIntVar(0, 10000, f'penalty_s{s}_{cat}_d{d+1}')
                        model.AddMaxEquality(nonneg, [temp, 0])
                        penalty_vars[s][cat] = nonneg

        model.Add(total_dispersion_penalty == sum(all_day_penalties) if all_day_penalties else 0)
        return total_dispersion_penalty
    def _create_solution_from_solver(self, solver, shifts, staff_list, day_list, fairness_group: set):
        schedule: Dict[datetime.date, List[Staff]] = {}
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

        counts = {staff.name: sum(solver.Value(shifts[(s_idx, d)]) for d in range(len(day_list)))
                  for s_idx, staff in enumerate(staff_list)}
        fairness_counts = self._calculate_fairness_group_counts(schedule, fairness_group)

        return {
            "schedule": schedule,
            "counts": counts,
            "fairness_group_counts": fairness_counts,
            "raw_shifts": raw_shifts_map,
        }

    def _calculate_fairness_group_counts(self, schedule: Dict[datetime.date, List[Staff]], fairness_group: set) -> Dict[str, int]:
        # Count only days that match categories in fairness_group (weekday names like '月'..'日' and/or '祝')
        if not fairness_group:
            return {s.name: 0 for s in self.all_staff}
        counts = {s.name: 0 for s in self.all_staff}
        for date_obj, staff_list in schedule.items():
            day_info = next((d for d in self.calendar_data if d['date'] == date_obj), None)
            if not day_info:
                continue
            is_holiday_selected = ('祝' in fairness_group and day_info.get('is_national_holiday', False))
            is_weekday_selected = (day_info.get('weekday') in fairness_group)
            if is_holiday_selected or is_weekday_selected:
                for staff in staff_list:
                    if staff.name in counts:
                        counts[staff.name] += 1
        return counts

    def _analyze_infeasibility(self, solver) -> str:
        """不充足時に、衝突している可能性が高い制約を簡易レポートとして返す。
        ORIGINE 相当の SufficientAssumptionsForInfeasibility 利用＋フォールバック。
        """
        try:
            assumptions = solver.SufficientAssumptionsForInfeasibility()
        except Exception:
            assumptions = []
        if not assumptions:
            # フォールバック: タグ一覧を付けた汎用メッセージ
            return (
                "シフトが見つかりませんでした。ルールを緩めて再試行してください。\n"
                "ヒント: 月限定休暇/固定、必要人数、休み間隔、最大連勤、特別日の公平性などを見直してください。"
            )
        lines = ["シフトが見つかりませんでした。以下のルールが衝突している可能性があります："]
        for idx in assumptions:
            tag = self.constraint_tags.get(idx)
            if tag:
                lines.append(f"・ {tag}")
            else:
                try:
                    lit = cp_model.Literal(idx)
                    tag = self.constraint_tags.get(lit.Var(), f"不明なルール({lit.Name()})")
                    lines.append(f"・ {tag}")
                except Exception:
                    pass
        return "\n".join(lines)


class SettingsManager:
    def __init__(self, history_dir: str = "shift_history"):
        self.staff_manager = StaffManager()
        self.rule_based_fixed_shifts: List[RuleBasedFixedShift] = []
        self.rule_based_vacations: List[RuleBasedVacation] = []
        self.min_interval = 2
        self.max_consecutive_days = 5
        self.shifts_per_day: dict = {'min': 1, 'max': 1}
        self.ignore_rules_on_holidays: bool = False
        self.avoid_consecutive_same_weekday: bool = False
        self.disperse_duties: bool = True
        self.fairness_group: set = set()
        self.max_solutions: int = 1
        self.fairness_tolerance: int = 1
        self.excel_title: str = "シフト表"
        self.history_dir = history_dir
        os.makedirs(self.history_dir, exist_ok=True)
        # 新規追加: 公平性のハード/ソフト切替とフォールバック
        self.fairness_as_hard: bool = True
        self.fallback_soft_on_infeasible: bool = True

    def to_dict(self) -> dict:
        staff_list_dict = [{
            "name": s.name,
            "color": s.color_code,
            "impossible_weekdays": sorted(list(s.impossible_weekdays)),
            "is_active": s.is_active,
        } for s in self.staff_manager.get_all_staff()]
        rules_fixed_dict = [{"week": r.week_number, "weekday": r.weekday_index, "staff_name": r.staff.name}
                            for r in self.rule_based_fixed_shifts]
        rules_vacation_dict = [{"week": r.week_number, "weekday": r.weekday_index, "staff_name": r.staff_name}
                               for r in self.rule_based_vacations]
        general_settings_dict = {
            "min_interval": self.min_interval,
            "max_consecutive_days": self.max_consecutive_days,
            "shifts_per_day": self.shifts_per_day,
            "ignore_rules_on_holidays": self.ignore_rules_on_holidays,
            "avoid_consecutive_same_weekday": self.avoid_consecutive_same_weekday,
            "disperse_duties": self.disperse_duties,
            "fairness_group": sorted(list(self.fairness_group)),
            "max_solutions": self.max_solutions,
            "fairness_tolerance": self.fairness_tolerance,
            "excel_title": self.excel_title,
            "fairness_as_hard": self.fairness_as_hard,
            "fallback_soft_on_infeasible": self.fallback_soft_on_infeasible,
        }
        return {"staff": staff_list_dict,
                "rule_based_fixed_shifts": rules_fixed_dict,
                "rule_based_vacations": rules_vacation_dict,
                "general_settings": general_settings_dict}

    @staticmethod
    def from_dict(data: dict) -> 'SettingsManager':
        settings = SettingsManager()
        staff_data = data.get("staff", [])
        for s in staff_data:
            color = s.get("color") or s.get("color_code") or "#000000"
            staff = Staff(s.get("name", ""), color, set(s.get("impossible_weekdays", [])), s.get("is_active", True))
            settings.staff_manager.add_or_update_staff(staff)
        rules_fixed_data = data.get("rule_based_fixed_shifts", [])
        for rf in rules_fixed_data:
            staff = settings.staff_manager.get_staff_by_name(rf.get("staff_name", ""))
            if staff:
                settings.rule_based_fixed_shifts.append(RuleBasedFixedShift(rf.get("week", 1), rf.get("weekday", 0), staff))
        rules_vacation_data = data.get("rule_based_vacations", [])
        for rv in rules_vacation_data:
            settings.rule_based_vacations.append(RuleBasedVacation(rv.get("week", 1), rv.get("weekday", 0), rv.get("staff_name", "")))
        general = data.get("general_settings", {})
        settings.min_interval = general.get("min_interval", 2)
        settings.max_consecutive_days = general.get("max_consecutive_days", 5)
        settings.shifts_per_day = general.get("shifts_per_day", {'min': 1, 'max': 1})
        settings.ignore_rules_on_holidays = general.get("ignore_rules_on_holidays", False)
        settings.avoid_consecutive_same_weekday = general.get("avoid_consecutive_same_weekday", False)
        settings.disperse_duties = general.get("disperse_duties", True)
        settings.fairness_group = set(general.get("fairness_group", []))
        settings.max_solutions = general.get("max_solutions", 1)
        settings.fairness_tolerance = general.get("fairness_tolerance", 1)
        settings.excel_title = general.get("excel_title", "シフト表")
        settings.fairness_as_hard = general.get("fairness_as_hard", True)
        settings.fallback_soft_on_infeasible = general.get("fallback_soft_on_infeasible", True)
        return settings

    def save_to_file(self, path: str) -> bool:
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(self.to_dict(), f, ensure_ascii=False, indent=2, default=str)
            return True
        except Exception:
            return False

    # Backward-compatible alias used by main.py
    def save_to_json(self, path: str) -> bool:
        return self.save_to_file(path)

    @staticmethod
    def load_from_file(path: str) -> Optional['SettingsManager']:
        try:
            with open(path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return SettingsManager.from_dict(data)
        except Exception:
            return None

    def save_history(self, year: int, month: int, solution: dict) -> bool:
        """履歴をJSONで保存する。
        フォーマットは ORIGINE に合わせ、schedule は
        [{"date": "YYYY-MM-DD", "staff_names": [..]}] の配列に正規化する。
        """
        try:
            # パス確保
            os.makedirs(self.history_dir, exist_ok=True)
            out_path = os.path.join(self.history_dir, f"{year:04d}-{month:02d}.json")

            # schedule 正規化
            schedule_for_json = []
            sched = solution.get("schedule", {})
            for date_obj, staff_list in sched.items():
                try:
                    date_str = date_obj.isoformat()
                except Exception:
                    # 既に文字列の可能性も考慮
                    date_str = str(date_obj)
                staff_names = []
                for s in (staff_list or []):
                    try:
                        staff_names.append(s.name)
                    except Exception:
                        staff_names.append(str(s))
                schedule_for_json.append({"date": date_str, "staff_names": staff_names})

            history_data = {
                "year": year,
                "month": month,
                "schedule": schedule_for_json,
                "counts": solution.get("counts", {}),
                "fairness_group_counts": solution.get("fairness_group_counts", {}),
            }

            with open(out_path, 'w', encoding='utf-8') as f:
                json.dump(history_data, f, ensure_ascii=False, indent=2)
            return True
        except Exception:
            return False

    def load_history(self, year: int, month: int) -> Optional[dict]:
        """履歴を読み込む。現行/旧来のファイル名どちらにも対応。"""
        try:
            # 現行パス
            path1 = os.path.join(self.history_dir, f"{year:04d}-{month:02d}.json")
            # 旧来（ORIGINE）パス
            path2 = os.path.join(self.history_dir, f"history_{year:04d}-{month:02d}.json")
            in_path = path1 if os.path.exists(path1) else (path2 if os.path.exists(path2) else None)
            if not in_path:
                return None
            with open(in_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return None

    def history_exists(self, year: int, month: int) -> bool:
        """履歴ファイルの存在を確認する。
        generation_tab.py から確認ダイアログの可否判断で利用される。
        """
        try:
            path1 = os.path.join(self.history_dir, f"{year:04d}-{month:02d}.json")
            path2 = os.path.join(self.history_dir, f"history_{year:04d}-{month:02d}.json")
            return os.path.exists(path1) or os.path.exists(path2)
        except Exception:
            return False




