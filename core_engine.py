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
            # 初回トライでも総回数の大きな偏りを避けるため、
            # 月限定固定などの強い制約がある場合は全体の総回数差分<=1をハード拘束
            if manual_fixed_shifts:
                try:
                    self._add_global_total_diff_leq1(model, shifts, staff_list, self.calendar_data)
                except Exception:
                    pass

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
                    # 通常生成が無理な場合のみ、段階的緩和を実施
                    try:
                        alt2 = self._progressive_relaxation(
                            shifts_per_day=shifts_per_day,
                            min_interval=min_interval,
                            max_consecutive_days=max_consecutive_days,
                            last_month_end_dates=last_month_end_dates,
                            prev_month_consecutive_days=prev_month_consecutive_days,
                            last_week_assignments=last_week_assignments,
                            avoid_consecutive_same_weekday=avoid_consecutive_same_weekday,
                            no_shift_dates=no_shift_dates,
                            manual_fixed_shifts=manual_fixed_shifts,
                            rule_based_fixed_shifts=rule_based_fixed_shifts,
                            vacations=vacations,
                            rule_based_vacations=rule_based_vacations,
                            fairness_group=fairness_group or set(),
                            total_adjustments=total_adjustments or {},
                            fairness_adjustments=fairness_adjustments or {},
                            fairness_tolerance=fairness_tolerance,
                            disperse_duties=disperse_duties,
                            past_schedules=self.past_schedules
                        )
                        if isinstance(alt2, list) and alt2:
                            return alt2
                        # 段階的緩和でも無理な場合、詳細付きの不可分析を返す
                        return alt2  # str
                    except Exception:
                        try:
                            return self._analyze_infeasibility(solver)
                        except Exception:
                            return "シフトが見つかりませんでした（制約の衝突）"
                break

        return found_solutions

    # ===== 段階的緩和ロジック =====
    def _progressive_relaxation(self,
                                *,
                                shifts_per_day,
                                min_interval: int,
                                max_consecutive_days: int,
                                last_month_end_dates,
                                prev_month_consecutive_days,
                                last_week_assignments,
                                avoid_consecutive_same_weekday: bool,
                                no_shift_dates,
                                manual_fixed_shifts,
                                rule_based_fixed_shifts,
                                vacations,
                                rule_based_vacations,
                                fairness_group: set,
                                total_adjustments: dict,
                                fairness_adjustments: dict,
                                fairness_tolerance: int,
                                disperse_duties: bool,
                                past_schedules: dict
                                ) -> List[dict] | str:
        """通常生成が無理な場合にのみ呼ばれる。段階1→段階2の順に最小限で緩和して解を得る。"""
        relaxations_log: List[str] = []
        # 段階1: 勤務間隔を 1 ずつ下げ、同時に元の間隔に対する近接違反をペナルティ化して再探索
        orig_min = max(0, int(min_interval))
        # 下限は 1 まで（ユーザ要件）
        for eff in range(orig_min - 1, 0, -1):
            sol = self._try_single_with_min_interval(
                min_interval_eff=eff,
                orig_min_interval=orig_min,
                shifts_per_day=shifts_per_day,
                max_consecutive_days=max_consecutive_days,
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
                past_schedules=past_schedules,
                enforce_total_diff_all_leq1=True,
                fairness_hard_names=None,
                fairness_soft_names=None
            )
            relaxations_log.append(f"勤務間隔を {orig_min}→{eff} に緩和（近接違反はペナルティ化）／総回数差分≤1をハード拘束")
            if isinstance(sol, dict):
                sol['relaxations'] = relaxations_log
                return [sol]
        # 1日でも無理なら段階2へ
        # 段階2: 厳しいスタッフを抽出し、期待回数（総回数）を個別に引き下げ
        tough_adj = self._estimate_tough_staff_adjustments(
            shifts_per_day, fairness_group, no_shift_dates, vacations, rule_based_vacations
        )
        if tough_adj:
            relaxations_log.append(
                "フェアネス: 問題スタッフの期待回数を個別に引き下げ (" + ", ".join(f"{k}:{v}" for k, v in tough_adj.items()) + ")"
            )
        # 合成調整（総回数の期待を tough だけ下げる）
        fa_total = dict(total_adjustments or {})
        for name, delta in tough_adj.items():
            fa_total[name] = fa_total.get(name, 0) - abs(int(delta))
        # Others の総回数レンジ計算（固定が多い人も tough に含める）
        staff_names = {st.name for st in self.all_staff}
        tough_names = set(tough_adj.keys())
        # 固定回数を集計し、ベースライン（総必要回数/人数）を超える人を tough に追加
        R_total = 0
        for day in self.calendar_data:
            mn, _ = self._get_shift_range_for_day(day, shifts_per_day)
            R_total += mn
        fixed_counts = self._count_manual_fixed_by_staff(manual_fixed_shifts)
        import math
        baseline = math.floor(R_total / max(1, len(self.all_staff)))
        tough_fixed = {name for name, cnt in fixed_counts.items() if cnt > baseline}
        if tough_fixed:
            relaxations_log.append("固定日数が多いスタッフを問題スタッフとして扱う: " + ", ".join(sorted(list(tough_fixed))))
        tough_names = tough_names | tough_fixed
        # 特別日/曜日フェアネスの期待も tough だけ下げる（固定超過や可用不足を反映）
        fa_fair = dict(fairness_adjustments or {})
        for name in tough_names:
            over_fixed = max(0, fixed_counts.get(name, 0) - baseline)
            delta_need = abs(int(tough_adj.get(name, 0))) + over_fixed
            if delta_need > 0:
                fa_fair[name] = fa_fair.get(name, 0) - delta_need
        if tough_names:
            relaxations_log.append("フェアネス: 問題スタッフの特別日/曜日期待回数を個別に引き下げ")
        others_names = staff_names - tough_names
        others_bounds = None
        if others_names:
            # 問題スタッフの月限定固定分（確定分）
            F_tough = 0
            for name, cnt in fixed_counts.items():
                if name in tough_names:
                    F_tough += cnt
            R_remain = max(0, R_total - F_tough)
            M = max(1, len(others_names))
            L = math.floor(R_remain / M)
            U = math.ceil(R_remain / M)
            others_bounds = (L, U)
            relaxations_log.append(f"非問題スタッフの総回数を範囲拘束: {L}〜{U}")
        # min_interval は 1 まで下げた状態＋近接違反ペナルティ＋Others範囲拘束で再探索
        sol2 = self._try_single_with_min_interval(
            min_interval_eff=1,
            orig_min_interval=orig_min,
            shifts_per_day=shifts_per_day,
            max_consecutive_days=max_consecutive_days,
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
            total_adjustments=fa_total,
            fairness_adjustments=fa_fair,
            fairness_tolerance=fairness_tolerance,
            disperse_duties=disperse_duties,
            past_schedules=past_schedules,
            others_total_bounds=others_bounds,
            others_names=others_names,
            fairness_hard_names=others_names,
            fairness_soft_names=tough_names
        )
        if isinstance(sol2, dict):
            sol2['relaxations'] = relaxations_log
            return [sol2]
        # 範囲拘束で不可なら、Others の総回数の差分 <= 1 のみに緩和して再試行
        if others_names:
            relaxations_log.append("非問題スタッフの総回数差分を1以内に緩和（範囲拘束を撤回）")
            sol3 = self._try_single_with_min_interval(
                min_interval_eff=1,
                orig_min_interval=orig_min,
                shifts_per_day=shifts_per_day,
                max_consecutive_days=max_consecutive_days,
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
                total_adjustments=fa_total,
                fairness_adjustments=fa_fair,
                fairness_tolerance=fairness_tolerance,
                disperse_duties=disperse_duties,
                past_schedules=past_schedules,
                others_names=others_names,
                enforce_others_diff_leq1=True,
                fairness_hard_names=others_names,
                fairness_soft_names=tough_names
            )
            if isinstance(sol3, dict):
                sol3['relaxations'] = relaxations_log
                return [sol3]
        # ここまでで解が出る想定。無理なら診断情報を付けて返す
        diag = self._build_stage2_diagnostic(
            tough_names=tough_names,
            fixed_counts=fixed_counts,
            R_total=R_total,
            others_names=others_names,
            others_bounds=others_bounds,
            max_consecutive_days=max_consecutive_days,
            manual_fixed_shifts=manual_fixed_shifts
        )
        base = sol2 if isinstance(sol2, str) else "シフトが見つかりませんでした（緩和後も不可）"
        return base + "\n\n" + diag

    def _try_single_with_min_interval(self,
                                      *,
                                      min_interval_eff: int,
                                      orig_min_interval: int,
                                      shifts_per_day,
                                      max_consecutive_days,
                                      last_month_end_dates,
                                      prev_month_consecutive_days,
                                      last_week_assignments,
                                      avoid_consecutive_same_weekday: bool,
                                      no_shift_dates,
                                      manual_fixed_shifts,
                                      rule_based_fixed_shifts,
                                      vacations,
                                      rule_based_vacations,
                                      fairness_group: set,
                                      total_adjustments: dict,
                                      fairness_adjustments: dict,
                                      fairness_tolerance: int,
                                      disperse_duties: bool,
                                      past_schedules: dict,
                                      others_total_bounds: Tuple[int, int] | None = None,
                                      others_names: Set[str] | None = None,
                                      enforce_others_diff_leq1: bool = False,
                                      enforce_total_diff_all_leq1: bool = False,
                                      fairness_hard_names: Set[str] | None = None,
                                      fairness_soft_names: Set[str] | None = None
                                      ) -> dict | str:
        model = cp_model.CpModel()
        self.constraint_tags = {}
        staff_list = self.all_staff
        day_list = self.calendar_data
        shifts = self._define_variables(model, staff_list, day_list)
        # ハード制約（min_interval は緩和後の値）
        self._add_hard_constraints(
            model, shifts, staff_list, day_list,
            no_shift_dates, shifts_per_day,
            rule_based_vacations, vacations,
            min_interval_eff, max_consecutive_days,
            last_month_end_dates, prev_month_consecutive_days,
            fairness_group, avoid_consecutive_same_weekday,
            last_week_assignments,
            manual_fixed_shifts,
            rule_based_fixed_shifts
        )
        # ソフト制約（ルール固定など）
        _, fixed_penalty = self._add_soft_constraints(model, shifts, staff_list, day_list,
                                                      rule_based_fixed_shifts, None)
        # 1B: 近接違反ペナルティ（元の min_interval を尊重）
        manual_fixed_lookup = self._manual_fixed_lookup(manual_fixed_shifts)
        minint_penalty = self._add_min_interval_soft_penalties(
            model, shifts, staff_list, day_list,
            orig_min_interval, manual_fixed_lookup
        )
        # Stage2/Stage1用: 総回数のハード拘束
        if (others_total_bounds or enforce_others_diff_leq1) and others_names:
            num_days = len(day_list)
            # 合計回数変数を各スタッフに作る
            total_vars: Dict[str, cp_model.IntVar] = {}
            for s_idx, st in enumerate(staff_list):
                v = model.NewIntVar(0, num_days, f'total_{s_idx}')
                model.Add(v == sum(shifts[(s_idx, d)] for d in range(num_days)))
                total_vars[st.name] = v
            selected_names = [n for n in sorted(others_names) if n in total_vars]
            selected = [total_vars[n] for n in selected_names]
            if selected:
                if others_total_bounds:
                    L, U = others_total_bounds
                    for name in selected_names:
                        var = total_vars[name]
                        cL = model.Add(var >= L)
                        cU = model.Add(var <= U)
                        try:
                            self.constraint_tags[cL.Index()] = f"{name}の総回数の下限（{L}）"
                            self.constraint_tags[cU.Index()] = f"{name}の総回数の上限（{U}）"
                        except Exception:
                            pass
                if enforce_others_diff_leq1 and len(selected) > 1:
                    min_v = model.NewIntVar(0, num_days, 'others_min_total')
                    max_v = model.NewIntVar(0, num_days, 'others_max_total')
                    model.AddMinEquality(min_v, selected)
                    model.AddMaxEquality(max_v, selected)
                    diff = model.NewIntVar(0, num_days, 'others_total_diff')
                    model.Add(diff == max_v - min_v)
                    cD = model.Add(diff <= 1)
                    try:
                        self.constraint_tags[cD.Index()] = "非問題スタッフの総回数差分≦1"
                    except Exception:
                        pass
        # 段階1対策: 全体の総回数差分<=1をハード拘束（要求A）
        if enforce_total_diff_all_leq1:
            num_days = len(day_list)
            all_totals = [model.NewIntVar(0, num_days, f'all_total_{s_idx}') for s_idx, _ in enumerate(staff_list)]
            for s_idx in range(len(staff_list)):
                model.Add(all_totals[s_idx] == sum(shifts[(s_idx, d)] for d in range(num_days)))
            min_all = model.NewIntVar(0, num_days, 'all_min_total')
            max_all = model.NewIntVar(0, num_days, 'all_max_total')
            model.AddMinEquality(min_all, all_totals)
            model.AddMaxEquality(max_all, all_totals)
            diff_all = model.NewIntVar(0, num_days, 'all_total_diff')
            model.Add(diff_all == max_all - min_all)
            cA = model.Add(diff_all <= 1)
            try:
                self.constraint_tags[cA.Index()] = "全体の総回数差分≦1（段階1）"
            except Exception:
                pass
        penalty_sum = self._sum_penalties(model, [fixed_penalty, minint_penalty])
        # 既存: 分散ペナルティ
        dispersion_penalty = 0
        if disperse_duties and fairness_group:
            dispersion_penalty = self._add_dispersion_penalty(
                model, shifts, staff_list, day_list, fairness_group, past_schedules
            )
        # 目的関数
        self._add_fairness_objective(
            model, shifts, staff_list, day_list,
            total_adjustments or {}, fairness_adjustments or {},
            fairness_tolerance, fairness_group or set(),
            penalty_sum, dispersion_penalty,
            hard_fair_staff_names=fairness_hard_names,
            soft_fair_staff_names=fairness_soft_names
        )
        solver = cp_model.CpSolver()
        status = solver.Solve(model)
        if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            sol = self._create_solution_from_solver(solver, shifts, staff_list, day_list, fairness_group or set())
            return sol
        if status == cp_model.INFEASIBLE:
            try:
                return self._analyze_infeasibility(solver)
            except Exception:
                return "シフトが見つかりませんでした（制約の衝突）"
        return "シフトが見つかりませんでした（未知の状態）"

    def _manual_fixed_lookup(self, manual_fixed_shifts: dict | None) -> Set[Tuple[str, datetime.date]]:
        result: Set[Tuple[str, datetime.date]] = set()
        if not manual_fixed_shifts:
            return result
        try:
            for d, lst in manual_fixed_shifts.items():
                for st in lst:
                    result.add((st.name, d))
        except Exception:
            pass
        return result

    def _add_min_interval_soft_penalties(self, model, shifts, staff_list, day_list,
                                          orig_min_interval: int,
                                          manual_fixed_lookup: Set[Tuple[str, datetime.date]]) -> cp_model.IntVar:
        """元の min_interval に対する近接違反をペナルティ化。r が小さいほど重く。
        月限定固定が両端のペアは免除。
        戻り値は合計ペナルティ用の IntVar。
        """
        num_days = len(day_list)
        penalty_lits: List[cp_model.IntVar] = []
        if orig_min_interval <= 1:
            zero = model.NewIntVar(0, 0, 'minint_penalty_zero')
            model.Add(zero == 0)
            return zero
        # 重み（距離1を重く）
        weights = {1: 100, 2: 60, 3: 40}
        for s_idx, staff in enumerate(staff_list):
            for d in range(num_days):
                for r in range(1, orig_min_interval):
                    k = d + r
                    if k >= num_days:
                        break
                    d_date = day_list[d]['date']
                    k_date = day_list[k]['date']
                    if (staff.name, d_date) in manual_fixed_lookup and (staff.name, k_date) in manual_fixed_lookup:
                        continue  # 月限定固定の連続は免除
                    x_d = shifts[(s_idx, d)]
                    x_k = shifts[(s_idx, k)]
                    v = model.NewBoolVar(f'v_close_{s_idx}_{d}_{r}')
                    model.Add(v <= x_d)
                    model.Add(v <= x_k)
                    model.Add(v >= x_d + x_k - 1)
                    # 重みを掛けた分だけ複製しなくても総和に反映するための IntVar を使う
                    # v は Bool のため、重み付きは以下のように表現する
                    w = weights.get(r, 40)
                    term = model.NewIntVar(0, w, f'v_close_w_{s_idx}_{d}_{r}')
                    model.Add(term == v * w)
                    penalty_lits.append(term)
        if penalty_lits:
            max_w = 100
            try:
                max_w = max(100, max([100, 60, 40]))
            except Exception:
                max_w = 100
            ub = len(penalty_lits) * max_w
            penalty_sum = model.NewIntVar(0, ub, 'minint_penalty')
            model.Add(penalty_sum == sum(penalty_lits))
            return penalty_sum
        zero = model.NewIntVar(0, 0, 'minint_penalty_zero2')
        model.Add(zero == 0)
        return zero

    def _sum_penalties(self, model, parts: List[int | cp_model.IntVar]) -> cp_model.IntVar:
        vals = []
        for p in parts:
            if isinstance(p, int):
                if p == 0:
                    continue
                c = model.NewIntVar(p, p, 'const_pen')
                vals.append(c)
            else:
                vals.append(p)
        if not vals:
            z = model.NewIntVar(0, 0, 'zero_pen')
            model.Add(z == 0)
            return z
        s = model.NewIntVar(0, 10_000_000, 'sum_pen')
        model.Add(s == sum(vals))
        return s

    def _add_global_total_diff_leq1(self, model, shifts, staff_list, day_list):
        num_days = len(day_list)
        totals = [model.NewIntVar(0, num_days, f'gtotal_{s}') for s in range(len(staff_list))]
        for s in range(len(staff_list)):
            model.Add(totals[s] == sum(shifts[(s, d)] for d in range(num_days)))
        min_v = model.NewIntVar(0, num_days, 'g_min_total')
        max_v = model.NewIntVar(0, num_days, 'g_max_total')
        model.AddMinEquality(min_v, totals)
        model.AddMaxEquality(max_v, totals)
        diff = model.NewIntVar(0, num_days, 'g_total_diff')
        model.Add(diff == max_v - min_v)
        c = model.Add(diff <= 1)
        try:
            self.constraint_tags[c.Index()] = "全体の総回数差分≦1（初回）"
        except Exception:
            pass

    def _count_manual_fixed_by_staff(self, manual_fixed_shifts: dict | None) -> Dict[str, int]:
        counts: Dict[str, int] = defaultdict(int)
        if not manual_fixed_shifts:
            return counts
        try:
            for _d, lst in manual_fixed_shifts.items():
                for st in lst:
                    counts[st.name] += 1
        except Exception:
            pass
        return counts

    def _build_stage2_diagnostic(self, *, tough_names: Set[str], fixed_counts: Dict[str, int], R_total: int,
                                  others_names: Set[str], others_bounds: Tuple[int, int] | None,
                                  max_consecutive_days: int,
                                  manual_fixed_shifts: dict | None) -> str:
        lines = []
        lines.append("[診断レポート: 段階2]")
        lines.append("")
        lines.append("■ 問題スタッフ（tough）")
        if tough_names:
            for name in sorted(list(tough_names)):
                lines.append(f"- {name}: 固定 {fixed_counts.get(name, 0)} 回")
        else:
            lines.append("- なし")
        lines.append("")
        lines.append(f"■ 総必要回数 R_total = {R_total}")
        lines.append("")
        lines.append("■ 非問題スタッフ（others）")
        if others_names:
            lines.append("- 対象: " + ", ".join(sorted(list(others_names))))
            if others_bounds:
                L, U = others_bounds
                lines.append(f"- 回数レンジ拘束: {L}〜{U}")
        else:
            lines.append("- なし")
        lines.append("")
        # 固定が最大連勤を超えていないか簡易チェック
        over_seq = []
        try:
            if manual_fixed_shifts:
                per_staff_dates: Dict[str, List[datetime.date]] = defaultdict(list)
                for d, lst in manual_fixed_shifts.items():
                    for st in lst:
                        per_staff_dates[st.name].append(d)
                for name, dlist in per_staff_dates.items():
                    dlist.sort()
                    run = 1
                    for i in range(1, len(dlist)):
                        if (dlist[i] - dlist[i-1]).days == 1:
                            run += 1
                            if run > max_consecutive_days:
                                over_seq.append((name, run))
                                break
                        else:
                            run = 1
        except Exception:
            pass
        lines.append("■ 固定と最大連勤の衝突")
        if over_seq:
            for name, run in over_seq:
                lines.append(f"- {name}: 固定が連続 {run} 日で最大連勤を超過")
        else:
            lines.append("- なし（固定による上限超過は検出されず）")
        return "\n".join(lines)

    def _estimate_tough_staff_adjustments(self, shifts_per_day, fairness_group: set,
                                          no_shift_dates, manual_vacations: dict | None,
                                          rule_based_vacations: List[RuleBasedVacation] | None) -> Dict[str, int]:
        """割当可能日が期待より著しく少ないスタッフを抽出し、その期待回数調整量（負値）を返す。"""
        staff_list = self.all_staff
        day_list = self.calendar_data
        # 期待総必要数（minの総和）をからおおまかに算出
        min_sum = 0
        for day in day_list:
            mn, _ = self._get_shift_range_for_day(day, shifts_per_day)
            min_sum += mn
        per_expected = max(0, round(min_sum / max(1, len(staff_list))))
        # スタッフごとの利用可能日数
        manual_vacations = manual_vacations or {}
        rb_vac_map = self._generate_vacations_from_rules(rule_based_vacations, day_list[0]['date'].year, day_list[0]['date'].month)
        tough: Dict[str, int] = {}
        for st in staff_list:
            avail = 0
            for day in day_list:
                date_obj = day['date']
                if no_shift_dates and date_obj in no_shift_dates:
                    continue
                if st.name in manual_vacations and date_obj in manual_vacations[st.name]:
                    continue
                if date_obj in (rb_vac_map.get(st.name, set())):
                    continue
                # 不可曜日。ただし祝日免除設定（ignore_rules_on_holidays）は solve 呼び出し側で保持しているが、
                # 本推定では厳しめに評価（祝日でも不可曜日は不可とみなす）
                if day['weekday'] in st.impossible_weekdays:
                    continue
                avail += 1
            deficit = per_expected - avail
            if deficit > 0:
                tough[st.name] = deficit
        return tough

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
                                 fixed_shift_penalty, dispersion_penalty,
                                 hard_fair_staff_names: Set[str] | None = None,
                                 soft_fair_staff_names: Set[str] | None = None):
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

                    # ハード評価対象の集合を決定（未指定時は従来どおり全員）
                    if hard_fair_staff_names is None and soft_fair_staff_names is None:
                        hard_indices = list(range(num_staff))
                        hard_label = "特別日回数の公平性 (許容差: {}回)".format(fairness_tolerance)
                        # 従来挙動
                        min_fair = model.NewIntVar(-num_days, num_days, 'min_fair')
                        max_fair = model.NewIntVar(0, num_days, 'max_fair')
                        model.AddMinEquality(min_fair, adj_fair)
                        model.AddMaxEquality(max_fair, adj_fair)
                        fair_diff = model.NewIntVar(0, num_days, 'fair_diff')
                        model.Add(fair_diff == max_fair - min_fair)
                        if self._fairness_as_hard:
                            c2 = model.Add(fair_diff <= fairness_tolerance)
                            try:
                                self.constraint_tags[c2.Index()] = hard_label
                            except Exception:
                                pass
                        else:
                            t = model.NewIntVar(-num_days, num_days, 'fair_over_tmp')
                            model.Add(t == fair_diff - fairness_tolerance)
                            over = model.NewIntVar(0, num_days, 'fair_over')
                            model.AddMaxEquality(over, [t, 0])
                            model.Add(fairness_penalty == over)
                    else:
                        # サブセット公平性: others をハード、tough はソフト（または無視）
                        name_at = [st.name for st in staff_list]
                        hard_indices = [i for i, nm in enumerate(name_at) if hard_fair_staff_names and nm in hard_fair_staff_names]
                        # ハード側（others）の公平性
                        if len(hard_indices) > 1 and self._fairness_as_hard:
                            hard_vars = [adj_fair[i] for i in hard_indices]
                            min_fair_h = model.NewIntVar(-num_days, num_days, 'min_fair_h')
                            max_fair_h = model.NewIntVar(0, num_days, 'max_fair_h')
                            model.AddMinEquality(min_fair_h, hard_vars)
                            model.AddMaxEquality(max_fair_h, hard_vars)
                            fair_diff_h = model.NewIntVar(0, num_days, 'fair_diff_h')
                            model.Add(fair_diff_h == max_fair_h - min_fair_h)
                            c2h = model.Add(fair_diff_h <= fairness_tolerance)
                            try:
                                self.constraint_tags[c2h.Index()] = f"特別日回数の公平性(othersのみ) (許容差: {fairness_tolerance}回)"
                            except Exception:
                                pass
                        # tough 側は今回はペナルティ無し（将来必要なら over を合算）

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




