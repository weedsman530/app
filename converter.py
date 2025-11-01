import pandas as pd
from tkinter import *
from tkinter import filedialog, ttk, messagebox
from datetime import datetime, time

def process_attendance():
    try:
        # اختيار الملفات
        dat_file = filedialog.askopenfilename(title="Select Attendance .dat File", filetypes=[("DAT Files", "*.dat"), ("Text Files", "*.txt")])
        emp_file = filedialog.askopenfilename(title="Select Employees File", filetypes=[("Excel Files", "*.xlsx")])
        if not dat_file or not emp_file:
            return
        
        # قراءة البيانات
        df = pd.read_csv(dat_file, sep=r"\s+", header=None, names=["ID","Date","Time","a","b","c","d"])
        df_emp = pd.read_excel(emp_file)

        df['Datetime'] = pd.to_datetime(df['Date'] + " " + df['Time'])
        df['DateOnly'] = pd.to_datetime(df['Date']).dt.date
        df['TimeOnly'] = pd.to_datetime(df['Time']).dt.time

        # ترتيب التواريخ من الأقدم للأحدث
        unique_dates = sorted(df["DateOnly"].unique())
        result = []

        # أوقات محددة
        target_in_start = time(7,30)
        target_in_end = time(8,30)
        breastfeeding_time = time(11,0)
        target_out_start = time(14,0)

        # معالجة كل يوم وكل موظف
        for day in unique_dates:
            df_day = df[df["DateOnly"] == day]
            for emp in df_emp.itertuples():
                emp_id = emp.code
                emp_name = emp.name
                
                emp_records = df_day[df_day["ID"] == emp_id]

                if emp_records.empty:
                    result.append([emp_id, day, emp_name, "غياب", "غياب"])
                    continue
                
                morning_records = emp_records[(emp_records['TimeOnly'] >= target_in_start) & (emp_records['TimeOnly'] <= target_in_end)]
                time_in = min(morning_records['TimeOnly']).strftime("%H:%M") if not morning_records.empty else "غياب"

                breastfeeding_records = emp_records[emp_records['TimeOnly'] == breastfeeding_time]
                if not breastfeeding_records.empty:
                    time_out = breastfeeding_time.strftime("%H:%M")
                else:
                    out_records = emp_records[emp_records['TimeOnly'] >= target_out_start]
                    time_out = max(out_records['TimeOnly']).strftime("%H:%M") if not out_records.empty else "غياب"

                result.append([emp_id, day, emp_name, time_in, time_out])

        # إنشاء الـDataFrame النهائي بدون عمود الملاحظات
        final_df = pd.DataFrame(result, columns=["Code","Date","Name","Time In","Time Out"])
        final_df["ملاحظات"] = ""

        # دالة الحفظ
        def save_file():
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")])
            if save_path:
                final_df.to_excel(save_path, index=False)
                messagebox.showinfo("✅ Done","Attendance sheet saved successfully!")

        # عرض النتائج في Treeview
        top = Toplevel(root)
        top.title("Attendance Results")
        top.geometry("1000x450")

        tree = ttk.Treeview(top, show='headings')
        tree.pack(fill=BOTH, expand=True)

        cols = list(final_df.columns)
        tree["columns"] = cols
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=150)

        for _, row in final_df.iterrows():
            tree.insert("", END, values=list(row))

        ttk.Button(top, text="💾 Save Excel", command=save_file).pack(pady=5)

    except Exception as e:
        messagebox.showerror("Error", str(e))


# الواجهة الرئيسية
root = Tk()
root.title("⏱ Attendance System")
root.geometry("400x250")

Label(root, text="📊 نظام حضور وانصراف ذكي", font=("Arial",14,"bold")).pack(pady=10)
ttk.Button(root, text="اختر الملفات وابدأ", command=process_attendance).pack(pady=20)

root.mainloop()
