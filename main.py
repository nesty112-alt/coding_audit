import sqlite3
import pandas as pd
import os
import json
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import warnings
from fuzzywuzzy import fuzz

# Pandas 경고 메시지 숨김
warnings.filterwarnings('ignore', category=UserWarning, module='pandas')

CONFIG_FILE = "config.json"

class ValidationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("데이터 검증 및 분석 프로그램")
        self.root.geometry("900x800")

        self.df_for_comparison = None
        self.comparison_result_df = None # 비교 결과 데이터프레임 저장
        self.loaded_filepath = ""

        self.config = {
            "main_csv": "", "main_rule": "", "sub_csv": "", "sub_rule": "",
            "output_dir": "", "compare_csv": ""
        }
        self.load_config()
        self.create_widgets()

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    self.config.update(json.load(f))
            except Exception: pass

    def save_config(self):
        self.config["main_csv"] = self.main_csv_var.get()
        self.config["main_rule"] = self.main_rule_var.get()
        self.config["sub_csv"] = self.sub_csv_var.get()
        self.config["sub_rule"] = self.sub_rule_var.get()
        self.config["output_dir"] = self.output_dir_var.get()
        
        path_to_save = self.loaded_filepath
        if not os.path.isfile(path_to_save):
            path_to_save = ""
        self.config["compare_csv"] = path_to_save

        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, ensure_ascii=False, indent=4)

    def _add_file_row(self, parent, label_text, var, row, is_dir=False):
        tk.Label(parent, text=label_text).grid(row=row, column=0, sticky="w", pady=2)
        tk.Entry(parent, textvariable=var, width=65).grid(row=row, column=1, padx=5, pady=2)
        def browse():
            path = filedialog.askdirectory() if is_dir else filedialog.askopenfilename(
                filetypes=[("CSV/Excel Files", "*.csv *.xlsx"), ("All Files", "*.*")])
            if path: var.set(path)
        tk.Button(parent, text="찾아보기", command=browse).grid(row=row, column=2, pady=2)

    def create_widgets(self):
        frame_common = tk.LabelFrame(self.root, text="공통 설정", padx=10, pady=10)
        frame_common.pack(fill="x", padx=10, pady=(10,0))
        self.output_dir_var = tk.StringVar(value=self.config.get("output_dir", ""))
        self._add_file_row(frame_common, "결과 저장 폴더:", self.output_dir_var, 0, is_dir=True)

        tab_control = ttk.Notebook(self.root)
        tab_sql = ttk.Frame(tab_control)
        tab_compare = ttk.Frame(tab_control)
        tab_control.add(tab_sql, text='   SQL 규칙 기반 검증   ')
        tab_control.add(tab_compare, text='   컬럼 간 텍스트 비교   ')
        tab_control.pack(expand=1, fill="both", padx=10, pady=10)

        self.create_sql_validation_tab(tab_sql)
        self.create_column_comparison_tab(tab_compare)

    def create_sql_validation_tab(self, parent_tab):
        # ... (이전과 동일)
        frame_paths = tk.LabelFrame(parent_tab, text="파일 경로 설정", padx=10, pady=10)
        frame_paths.pack(fill="x", padx=10, pady=5)
        self.main_csv_var = tk.StringVar(value=self.config.get("main_csv", ""))
        self.main_rule_var = tk.StringVar(value=self.config.get("main_rule", ""))
        self.sub_csv_var = tk.StringVar(value=self.config.get("sub_csv", ""))
        self.sub_rule_var = tk.StringVar(value=self.config.get("sub_rule", ""))
        for var in (self.main_csv_var, self.main_rule_var, self.sub_csv_var, self.sub_rule_var):
            var.trace_add("write", self.auto_select_run_option)
        self._add_file_row(frame_paths, "Main 데이터 (CSV):", self.main_csv_var, 0)
        self._add_file_row(frame_paths, "Main 검증 룰 (Excel):", self.main_rule_var, 1)
        tk.Label(frame_paths, text="-" * 110).grid(row=2, column=0, columnspan=3)
        self._add_file_row(frame_paths, "Sub 데이터 (CSV):", self.sub_csv_var, 3)
        self._add_file_row(frame_paths, "Sub 검증 룰 (Excel):", self.sub_rule_var, 4)

        frame_options = tk.Frame(parent_tab)
        frame_options.pack(fill="x", padx=10, pady=10)
        self.run_option_var = tk.IntVar(value=3)
        tk.Radiobutton(frame_options, text="Main만 실행", variable=self.run_option_var, value=1).pack(side="left", padx=10)
        tk.Radiobutton(frame_options, text="Sub만 실행", variable=self.run_option_var, value=2).pack(side="left", padx=10)
        tk.Radiobutton(frame_options, text="모두 실행", variable=self.run_option_var, value=3).pack(side="left", padx=10)
        self.run_btn_text = tk.StringVar(value="▶ SQL 검증 실행")
        self.run_btn = tk.Button(frame_options, textvariable=self.run_btn_text, bg="#4CAF50", fg="white", font=("Arial", 11, "bold"), command=self.run_process)
        self.run_btn.pack(side="right", padx=10)
        self.auto_select_run_option()

        frame_results = tk.LabelFrame(parent_tab, text="검증 결과", padx=10, pady=10)
        frame_results.pack(fill="both", expand=True, padx=10, pady=5)
        columns = ("target", "rule_name", "status", "total_cnt", "error_cnt", "error_msg")
        self.tree = ttk.Treeview(frame_results, columns=columns, show="headings")
        self.tree.heading("target", text="타겟"); self.tree.column("target", width=60, anchor="center")
        self.tree.heading("rule_name", text="검증 룰"); self.tree.column("rule_name", width=220, anchor="w")
        self.tree.heading("status", text="결과"); self.tree.column("status", width=80, anchor="center")
        self.tree.heading("total_cnt", text="총 데이터"); self.tree.column("total_cnt", width=90, anchor="e")
        self.tree.heading("error_cnt", text="오류 건수"); self.tree.column("error_cnt", width=80, anchor="e")
        self.tree.heading("error_msg", text="비고"); self.tree.column("error_msg", width=300, anchor="w")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", self.on_tree_double_click)
        self.tree.bind("<Button-3>", self.show_context_menu)


    def create_column_comparison_tab(self, parent_tab):
        # --- 설정 영역 ---
        frame_compare = tk.LabelFrame(parent_tab, text="비교 설정 (단일 파일)", padx=10, pady=10)
        frame_compare.pack(fill="x", padx=10, pady=10)
        frame_compare.columnconfigure(1, weight=1)

        self.loaded_filepath = self.config.get("compare_csv", "")
        display_text = self.loaded_filepath if self.loaded_filepath else "선택된 파일 없음"
        self.compare_csv_label_var = tk.StringVar(value=display_text)

        tk.Button(frame_compare, text="비교 대상 파일 불러오기", command=self.select_and_load_comparison_file).grid(row=0, column=0, padx=5, pady=5, sticky="w")
        tk.Label(frame_compare, textvariable=self.compare_csv_label_var, fg="blue", wraplength=500, justify="left").grid(row=0, column=1, sticky="w", padx=5)

        self.compare_btn = tk.Button(frame_compare, text="▶ 비교 실행", bg="#FF9800", fg="white", font=("Arial", 10, "bold"), command=self.run_column_comparison)
        self.compare_btn.grid(row=0, column=2, padx=20, ipady=10, sticky="e")

        # --- Listbox 영역 ---
        list_frame = tk.Frame(parent_tab)
        list_frame.pack(fill="both", expand=True, padx=10, pady=10)
        list_frame.columnconfigure(0, weight=1)
        list_frame.columnconfigure(1, weight=1)
        list_frame.rowconfigure(1, weight=1)

        tk.Label(list_frame, text="컬럼 1 선택:").grid(row=0, column=0, sticky="w", padx=5)
        list1_subframe = tk.Frame(list_frame)
        list1_subframe.grid(row=1, column=0, sticky="nsew", padx=5)
        list1_subframe.rowconfigure(0, weight=1)
        list1_subframe.columnconfigure(0, weight=1)
        scrollbar1 = tk.Scrollbar(list1_subframe)
        scrollbar1.grid(row=0, column=1, sticky="ns")
        self.col1_listbox = tk.Listbox(list1_subframe, yscrollcommand=scrollbar1.set, exportselection=False)
        self.col1_listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar1.config(command=self.col1_listbox.yview)

        tk.Label(list_frame, text="컬럼 2 선택:").grid(row=0, column=1, sticky="w", padx=5)
        list2_subframe = tk.Frame(list_frame)
        list2_subframe.grid(row=1, column=1, sticky="nsew", padx=5)
        list2_subframe.rowconfigure(0, weight=1)
        list2_subframe.columnconfigure(0, weight=1)
        scrollbar2 = tk.Scrollbar(list2_subframe)
        scrollbar2.grid(row=0, column=1, sticky="ns")
        self.col2_listbox = tk.Listbox(list2_subframe, yscrollcommand=scrollbar2.set, exportselection=False)
        self.col2_listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar2.config(command=self.col2_listbox.yview)

        # --- 결과 요약 및 저장 영역 ---
        summary_frame = tk.LabelFrame(parent_tab, text="비교 결과 요약", padx=10, pady=10)
        summary_frame.pack(fill="x", padx=10, pady=5)
        
        self.summary_total_var = tk.StringVar(value="N/A")
        self.summary_avg_var = tk.StringVar(value="N/A")
        self.summary_exact_var = tk.StringVar(value="N/A")
        self.summary_partial_var = tk.StringVar(value="N/A")

        tk.Label(summary_frame, text="총 비교 건수:").grid(row=0, column=0, sticky="w", padx=5, pady=2)
        tk.Label(summary_frame, textvariable=self.summary_total_var, font=("Arial", 10, "bold")).grid(row=0, column=1, sticky="w", padx=5)
        tk.Label(summary_frame, text="평균 유사도:").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        tk.Label(summary_frame, textvariable=self.summary_avg_var, font=("Arial", 10, "bold")).grid(row=1, column=1, sticky="w", padx=5)
        tk.Label(summary_frame, text="완전 일치 (100점):").grid(row=0, column=2, sticky="w", padx=20, pady=2)
        tk.Label(summary_frame, textvariable=self.summary_exact_var, font=("Arial", 10, "bold")).grid(row=0, column=3, sticky="w", padx=5)
        tk.Label(summary_frame, text="부분 일치 (90점 이상):").grid(row=1, column=2, sticky="w", padx=20, pady=2)
        tk.Label(summary_frame, textvariable=self.summary_partial_var, font=("Arial", 10, "bold")).grid(row=1, column=3, sticky="w", padx=5)

        self.save_btn = tk.Button(summary_frame, text="결과 저장하기", bg="#2196F3", fg="white", state="disabled", command=self.save_comparison_results)
        self.save_btn.grid(row=0, column=4, rowspan=2, padx=30, ipady=10, sticky="e")
        summary_frame.columnconfigure(4, weight=1)


    def select_and_load_comparison_file(self):
        csv_path = filedialog.askopenfilename(
            title="비교할 CSV 파일을 선택하세요",
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")])
        if not csv_path: return

        self.loaded_filepath = csv_path
        self.compare_csv_label_var.set(csv_path)
        
        self.col1_listbox.delete(0, tk.END)
        self.col2_listbox.delete(0, tk.END)
        self._reset_summary()

        try:
            try: self.df_for_comparison = pd.read_csv(csv_path, encoding='utf-8', low_memory=False)
            except UnicodeDecodeError: self.df_for_comparison = pd.read_csv(csv_path, encoding='cp949', low_memory=False)
            
            columns = sorted(self.df_for_comparison.columns.tolist())
            for col in columns:
                self.col1_listbox.insert(tk.END, col)
                self.col2_listbox.insert(tk.END, col)
            
            messagebox.showinfo("성공", "파일을 성공적으로 불러왔습니다.\n이제 비교할 컬럼을 선택해주세요.")
        except Exception as e:
            messagebox.showerror("파일 로드 오류", f"파일을 읽는 중 오류가 발생했습니다:\n{e}")
            self.df_for_comparison = None
            self.compare_csv_label_var.set("파일 로드 실패. 다시 시도해주세요.")
            self.loaded_filepath = ""

    def _reset_summary(self):
        self.summary_total_var.set("N/A")
        self.summary_avg_var.set("N/A")
        self.summary_exact_var.set("N/A")
        self.summary_partial_var.set("N/A")
        self.save_btn.config(state="disabled")
        self.comparison_result_df = None

    def run_column_comparison(self):
        col1_idx = self.col1_listbox.curselection()
        col2_idx = self.col2_listbox.curselection()

        if self.df_for_comparison is None:
            messagebox.showerror("오류", "먼저 '비교 대상 파일 불러오기'를 실행하여 데이터를 로드해주세요.")
            return
        if not col1_idx or not col2_idx:
            messagebox.showerror("입력 오류", "비교할 두 개의 컬럼을 각각 하나씩 선택해야 합니다.")
            return

        col1 = self.col1_listbox.get(col1_idx[0])
        col2 = self.col2_listbox.get(col2_idx[0])

        if col1 == col2:
            messagebox.showwarning("경고", "서로 다른 두 개의 컬럼을 선택해주세요.")
            return

        self._reset_summary()
        self.compare_btn.config(state="disabled", text="⏳ 비교 중...")
        thread = threading.Thread(target=self._run_comparison_thread, args=(col1, col2))
        thread.daemon = True
        thread.start()

    def _run_comparison_thread(self, col1, col2):
        try:
            # 원본 데이터프레임 전체를 복사하여 사용
            df = self.df_for_comparison.copy()
            
            # 유사도 계산 (NaN 값은 빈 문자열로 처리)
            df['similarity_score'] = df.apply(
                lambda row: fuzz.ratio(
                    str(row[col1]) if pd.notna(row[col1]) else '',
                    str(row[col2]) if pd.notna(row[col2]) else ''
                ),
                axis=1
            )
            
            self.comparison_result_df = df

            # 요약 정보 계산
            total_rows = len(self.comparison_result_df)
            avg_score = self.comparison_result_df['similarity_score'].mean()
            exact_matches = (self.comparison_result_df['similarity_score'] == 100).sum()
            partial_matches = (self.comparison_result_df['similarity_score'] >= 90).sum()

            summary_data = {
                "total": f"{total_rows:,} 건",
                "avg": f"{avg_score:.2f} 점",
                "exact": f"{exact_matches:,} 건",
                "partial": f"{partial_matches:,} 건"
            }
            self.root.after(0, self._update_summary_ui, summary_data)

        except Exception as e:
            self.root.after(0, lambda e=e: messagebox.showerror("비교 오류", f"비교 작업 중 오류가 발생했습니다:\n{str(e)}"))
        finally:
            self.root.after(0, lambda: self.compare_btn.config(state="normal", text="▶ 비교 실행"))

    def _update_summary_ui(self, summary_data):
        self.summary_total_var.set(summary_data["total"])
        self.summary_avg_var.set(summary_data["avg"])
        self.summary_exact_var.set(summary_data["exact"])
        self.summary_partial_var.set(summary_data["partial"])
        self.save_btn.config(state="normal")
        messagebox.showinfo("비교 완료", "텍스트 비교 분석이 완료되었습니다. 요약 정보를 확인하고 결과를 저장할 수 있습니다.")

    def save_comparison_results(self):
        out_dir = self.output_dir_var.get().strip()
        if not out_dir:
            messagebox.showerror("오류", "먼저 상단의 '공통 설정'에서 결과 저장 폴더를 지정해주세요.")
            return
        if self.comparison_result_df is None:
            messagebox.showerror("오류", "저장할 비교 결과가 없습니다. 먼저 '비교 실행'을 수행해주세요.")
            return
        
        try:
            # 컬럼 순서 재정렬 (similarity_score를 맨 앞으로)
            cols = self.comparison_result_df.columns.tolist()
            cols.insert(0, cols.pop(cols.index('similarity_score')))
            df_to_save = self.comparison_result_df[cols]

            output_path = os.path.join(out_dir, "column_comparison_result.csv")
            df_to_save.to_csv(output_path, index=False, encoding='utf-8-sig')
            messagebox.showinfo("저장 완료", f"비교 결과가 성공적으로 저장되었습니다.\n경로: {output_path}")
        except Exception as e:
            messagebox.showerror("저장 오류", f"파일 저장 중 오류가 발생했습니다:\n{e}")

    # --- SQL 검증 관련 메소드들 ---
    def auto_select_run_option(self, *args):
        # ... (이전과 동일)
        main_ready = bool(self.main_csv_var.get().strip() and self.main_rule_var.get().strip())
        sub_ready = bool(self.sub_csv_var.get().strip() and self.sub_rule_var.get().strip())
        if main_ready and sub_ready: self.run_option_var.set(3)
        elif main_ready: self.run_option_var.set(1)
        elif sub_ready: self.run_option_var.set(2)

    def show_context_menu(self, event):
        # ... (이전과 동일)
        row_id = self.tree.identify_row(event.y)
        if not row_id: return
        self.tree.selection_set(row_id)
        self.tree.focus(row_id)
        item_values = self.tree.item(row_id, "values")
        if self.tree.identify_column(event.x) == '#6':
            error_msg = item_values[5]
            if error_msg and error_msg != "-":
                menu = tk.Menu(self.root, tearoff=0)
                menu.add_command(label="비고 내용 복사", command=lambda: self.copy_to_clipboard_direct(error_msg))
                menu.post(event.x_root, event.y_root)

    def copy_to_clipboard_direct(self, text):
        # ... (이전과 동일)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)

    def on_tree_double_click(self, event):
        # ... (이전과 동일)
        item = self.tree.selection()
        if not item: return
        error_msg = self.tree.item(item[0], "values")[5]
        if error_msg and error_msg != "-":
            self.show_error_popup(self.tree.item(item[0], "values")[1], error_msg)

    def show_error_popup(self, rule_name, message):
        # ... (이전과 동일)
        popup = tk.Toplevel(self.root)
        popup.title("상세 오류 메시지")
        popup.geometry("600x300")
        popup.attributes('-topmost', True)
        tk.Label(popup, text=f"검증 룰: {rule_name}", font=("Arial", 10, "bold")).pack(anchor="w", padx=10, pady=5)
        text = tk.Text(popup, wrap="word", font=("Consolas", 10))
        text.insert("1.0", message)
        text.config(state="disabled")
        text.pack(fill="both", expand=True, padx=10, pady=10)
        def copy():
            self.root.clipboard_clear()
            self.root.clipboard_append(message)
            btn.config(text="✓ 복사 완료")
        btn = tk.Button(popup, text="메시지 복사", command=copy)
        btn.pack(pady=5)

    def run_process(self):
        # ... (이전과 동일)
        opt = self.run_option_var.get()
        out_dir = self.output_dir_var.get().strip()
        if not out_dir:
            messagebox.showerror("오류", "결과 저장 폴더를 지정해주세요.")
            return
        main_c = self.main_csv_var.get().strip()
        main_r = self.main_rule_var.get().strip()
        sub_c = self.sub_csv_var.get().strip()
        sub_r = self.sub_rule_var.get().strip()
        if opt == 1 and (not main_c or not main_r):
            messagebox.showwarning("경로 누락", "Main 실행을 위해 데이터와 룰 파일을 모두 입력해주세요.")
            return
        elif opt == 2 and (not sub_c or not sub_r):
            messagebox.showwarning("경로 누락", "Sub 실행을 위해 데이터와 룰 파일을 모두 입력해주세요.")
            return
        elif opt == 3 and not all([main_c, main_r, sub_c, sub_r]):
            messagebox.showwarning("경로 누락", "모두 실행하려면 4개 파일 경로를 전부 입력해야 합니다.")
            return
        self.save_config()
        for item in self.tree.get_children(): self.tree.delete(item)
        self.run_btn.config(state="disabled")
        self.run_btn_text.set("⏳ 실행 중...")
        thread = threading.Thread(target=self._run_process_thread, args=(opt, out_dir))
        thread.daemon = True
        thread.start()

    def _safe_tree_insert(self, values):
        # ... (이전과 동일)
        self.tree.insert("", "end", values=values)

    def _run_process_thread(self, opt, out_dir):
        # ... (이전과 동일)
        conn = sqlite3.connect(':memory:')
        try:
            if opt in (1, 3):
                self._load_and_validate("Main", self.main_csv_var.get(), self.main_rule_var.get(), out_dir, conn)
            if opt in (2, 3):
                self._load_and_validate("Sub", self.sub_csv_var.get(), self.sub_rule_var.get(), out_dir, conn)
        except Exception as e:
            self.root.after(0, lambda e=e: messagebox.showerror("실행 오류", f"작업 중 오류가 발생했습니다:\n{str(e)}"))
        finally:
            conn.close()
            self.root.after(0, lambda: self.run_btn.config(state="normal"))
            self.root.after(0, lambda: self.run_btn_text.set("▶ SQL 검증 실행"))

    def _load_and_validate(self, target_name, csv_path, rule_path, out_dir, conn):
        # ... (이전과 동일)
        if not csv_path or not rule_path: return
        table_name = target_name.lower()
        try:
            try: df_data = pd.read_csv(csv_path, encoding='utf-8', low_memory=False)
            except: df_data = pd.read_csv(csv_path, encoding='cp949', low_memory=False)
            df_data.to_sql(table_name, conn, if_exists='replace', index=False)
            total_records = len(df_data)
        except Exception as e: raise Exception(f"{target_name} 데이터 로드 실패: {e}")
        try:
            df_rules = pd.read_excel(rule_path)
        except Exception as e: raise Exception(f"{target_name} 룰 파일 로드 실패: {e}")

        all_errors = []
        output_excel_path = os.path.join(out_dir, f"{table_name}_audit_result.xlsx")
        for _, row in df_rules.iterrows():
            rule_id = str(row.get('number', ''))
            rule_name = str(row.get('Rule_Name', ''))
            sql_query = row.get('SQL_Query', '')
            if not sql_query or pd.isna(sql_query): continue
            display_name = f"[{rule_id}] {rule_name}" if rule_id else rule_name
            try:
                result_df = pd.read_sql_query(sql_query, conn)
                error_count = len(result_df)
                status = "🚨 오류" if error_count > 0 else "✅ 정상"
                self.root.after(0, self._safe_tree_insert, (target_name, display_name, status, f"{total_records:,}건", f"{error_count:,}건", "-"))
                if not result_df.empty:
                    temp_df = result_df.copy()
                    temp_df.insert(0, '오류유형', display_name)
                    all_errors.append(temp_df)
            except Exception as e:
                error_msg = str(e).replace('\n', ' ')
                self.root.after(0, self._safe_tree_insert, (target_name, display_name, "⚠️ SQL실패", f"{total_records:,}건", "N/A", error_msg))
        
        if all_errors:
            final_error_df = pd.concat(all_errors, ignore_index=True)
            with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
                final_error_df.to_excel(writer, sheet_name='오류내역', index=False)
        else:
            pd.DataFrame([{'결과': '발견된 오류 데이터가 없습니다.'}]).to_excel(output_excel_path, index=False)

if __name__ == "__main__":
    root = tk.Tk()
    app = ValidationApp(root)
    root.mainloop()
