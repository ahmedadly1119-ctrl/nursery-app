import streamlit as st
import pandas as pd
import datetime
import calendar
import os
from io import BytesIO

# --- 1. إعدادات قاعدة البيانات (Excel) ---
DB_FILE = "nursery_data.xlsx"

def to_excel(df):
    """وظيفة تحويل البيانات لملف إكسيل قابل للتحميل"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

def load_data():
    if os.path.exists(DB_FILE):
        try:
            students = pd.read_excel(DB_FILE, sheet_name='Students')
            ledger = pd.read_excel(DB_FILE, sheet_name='Ledger')
            students['تاريخ التسجيل/Reg Date'] = pd.to_datetime(students['تاريخ التسجيل/Reg Date']).dt.date
            ledger['التاريخ/Date'] = pd.to_datetime(ledger['التاريخ/Date'])
            return students, ledger
        except Exception:
            return create_empty_db()
    else:
        return create_empty_db()

def create_empty_db():
    students = pd.DataFrame(columns=['ID', 'الاسم/Name', 'ولي الأمر/Parent', 'الهاتف/Phone', 'المبلغ الثابت/Fee', 'الحالة/Status', 'تاريخ التسجيل/Reg Date', 'الرصيد/Balance'])
    ledger = pd.DataFrame(columns=['EntryID', 'التاريخ/Date', 'اسم الطالب/Student', 'النوع/Type', 'البيان/Desc', 'مدين/Debit (+)', 'دائن/Credit (-)', 'الرصيد بعد الحركة/Balance'])
    return students, ledger

def save_to_excel(students, ledger):
    try:
        with pd.ExcelWriter(DB_FILE, engine='openpyxl') as writer:
            students.to_excel(writer, sheet_name='Students', index=False)
            ledger.to_excel(writer, sheet_name='Ledger', index=False)
    except Exception as e:
        st.error(f"⚠️ تنبيه: الملف مفتوح في إكسيل، اغلقه ليتم الحفظ تلقائياً.")

# --- 2. إعدادات الصفحة واللغة ---
st.set_page_config(page_title="Nursery Management PRO V11.3", layout="wide")

if 'students' not in st.session_state or 'ledger' not in st.session_state:
    st.session_state.students, st.session_state.ledger = load_data()

if 'lang' not in st.session_state:
    st.session_state.lang = 'Arabic'

translations = {
    'Arabic': {
        'title': "🏫 نظام إدارة الحضانة الذكي",
        'menu': ["الرئيسية 📊", "ملفات الطلاب 📂", "الفوترة والتسويات ⏰", "تحصيل النقدية 💵", "التقارير والحسابات 📄", "تقارير الفترة 📥"],
        'total_debt': "إجمالي الديون القائمة", 'total_paid': "إجمالي المحصل فعلياً", 'active_st': "الطلاب النشطين",
        'st_name': "اسم الطالب", 'parent': "ولي الأمر", 'phone': "رقم الهاتف", 'fee': "الاشتراك الشهري",
        'reg_date': "تاريخ الالتحاق", 'save': "حفظ البيانات", 'update': "تحديث البيانات", 'status': "الحالة",
        'monthly_fees': "مديونية الشهر الجماعية", 'extra_hours': "ساعات إضافية", 'adjustments': "تسويات وخصومات",
        'download': "تحميل ملف إكسيل", 'delete_last': "❌ حذف آخر حركة (تراجع)", 'confirm': "تأكيد وتنفيذ",
        'summary_rep': "كشف الأرصدة المجمع", 'detailed_rep': "كشف حساب تفصيلي"
    },
    'English': {
        'title': "🏫 Smart Nursery PRO",
        'menu': ["Dashboard 📊", "Student Files 📂", "Billing & Adj ⏰", "Cash Collection 💵", "Reports & Ledger 📄", "Period Reports 📥"],
        'total_debt': "Total Debt", 'total_paid': "Total Collected", 'active_st': "Active Students",
        'st_name': "Student Name", 'parent': "Parent Name", 'phone': "Phone Number", 'fee': "Monthly Fee",
        'reg_date': "Join Date", 'save': "Save Data", 'update': "Update Data", 'status': "Status",
        'monthly_fees': "Bulk Monthly Fees", 'extra_hours': "Extra Hours", 'adjustments': "Adjustments",
        'download': "Download Excel", 'delete_last': "❌ Undo Last Transaction", 'confirm': "Confirm",
        'summary_rep': "Balances Summary", 'detailed_rep': "Detailed Ledger"
    }
}

st.sidebar.title("Settings / الإعدادات")
st.session_state.lang = st.sidebar.selectbox("Language / اللغة", ["Arabic", "English"])
T = translations[st.session_state.lang]

# --- 3. محرك العمليات ---
def record_transaction(name, t_type, desc, amount, custom_date=None):
    idx = st.session_state.students[st.session_state.students['الاسم/Name'] == name].index[0]
    curr_bal = st.session_state.students.loc[idx, 'الرصيد/Balance']
    debit = amount if t_type in ["Invoice", "Debit"] else 0
    credit = amount if t_type in ["Payment", "Discount"] else 0
    new_bal = curr_bal + debit - credit
    st.session_state.students.loc[idx, 'الرصيد/Balance'] = new_bal
    eid = st.session_state.ledger['EntryID'].max() + 1 if not st.session_state.ledger.empty else 1
    dt_final = datetime.datetime.combine(custom_date if custom_date else datetime.date.today(), datetime.time.min)
    new_e = {'EntryID': eid, 'التاريخ/Date': dt_final, 'اسم الطالب/Student': name, 'النوع/Type': t_type, 'البيان/Desc': desc, 'مدين/Debit (+)': debit, 'دائن/Credit (-)': credit, 'الرصيد بعد الحركة/Balance': new_bal}
    st.session_state.ledger = pd.concat([st.session_state.ledger, pd.DataFrame([new_e])], ignore_index=True)
    save_to_excel(st.session_state.students, st.session_state.ledger)

# --- 4. القائمة والواجهات ---
menu = st.sidebar.radio("Navigate", T['menu'])

# 4.1 الرئيسية
if menu == T['menu'][0]:
    st.header(T['menu'][0])
    c1, c2, c3 = st.columns(3)
    c1.metric(T['total_debt'], f"{st.session_state.students['الرصيد/Balance'].sum():,.2f}")
    c2.metric(T['total_paid'], f"{st.session_state.ledger['دائن/Credit (-)'].sum():,.2f}")
    c3.metric(T['active_st'], len(st.session_state.students[st.session_state.students['الحالة/Status']=='نشط']))
    st.divider()
    late = st.session_state.students[st.session_state.students['الرصيد/Balance'] > 0]
    if not late.empty:
        st.subheader("🔔 Overdue")
        st.dataframe(late[['الاسم/Name', 'الهاتف/Phone', 'الرصيد/Balance']], use_container_width=True)

# 4.2 ملفات الطلاب
elif menu == T['menu'][1]:
    t1, t2 = st.tabs(["🆕 Add Student", "⚙️ Edit Profile"])
    with t1:
        with st.form("add_form"):
            col1, col2 = st.columns(2)
            n = col1.text_input(T['st_name']); p = col2.text_input(T['parent'])
            ph = col1.text_input(T['phone']); f = col2.number_input(T['fee'], min_value=0)
            rd = col1.date_input(T['reg_date'], datetime.date.today())
            if st.form_submit_button(T['save']):
                if n:
                    new_s = {'ID': 101+len(st.session_state.students), 'الاسم/Name': n, 'ولي الأمر/Parent': p, 'الهاتف/Phone': ph, 'المبلغ الثابت/Fee': f, 'الحالة/Status': 'نشط', 'تاريخ التسجيل/Reg Date': rd, 'الرصيد/Balance': 0}
                    st.session_state.students = pd.concat([st.session_state.students, pd.DataFrame([new_s])], ignore_index=True)
                    save_to_excel(st.session_state.students, st.session_state.ledger); st.success("Saved!")
    with t2:
        if not st.session_state.students.empty:
            target = st.selectbox(T['st_name'], st.session_state.students['الاسم/Name'])
            idx = st.session_state.students[st.session_state.students['الاسم/Name'] == target].index[0]
            curr = st.session_state.students.iloc[idx]
            with st.form("edit_form"):
                col1, col2 = st.columns(2)
                u_n = col1.text_input(T['st_name'], curr['الاسم/Name'])
                u_p = col2.text_input(T['parent'], curr['ولي الأمر/Parent'])
                u_ph = col1.text_input(T['phone'], curr['الهاتف/Phone'])
                u_f = col2.number_input(T['fee'], value=int(curr['المبلغ الثابت/Fee']))
                u_bal = col1.number_input("Manual Balance Adjust", value=float(curr['الرصيد/Balance']))
                u_st = col2.selectbox(T['status'], ["نشط", "مجمد", "منقطع"], index=0)
                u_rd = col1.date_input(T['reg_date'], value=curr['تاريخ التسجيل/Reg Date'])
                if st.form_submit_button(T['update']):
                    st.session_state.students.at[idx, 'الاسم/Name'] = u_n
                    st.session_state.students.at[idx, 'ولي الأمر/Parent'] = u_p
                    st.session_state.students.at[idx, 'الهاتف/Phone'] = u_ph
                    st.session_state.students.at[idx, 'المبلغ الثابت/Fee'] = u_f
                    st.session_state.students.at[idx, 'الرصيد/Balance'] = u_bal
                    st.session_state.students.at[idx, 'الحالة/Status'] = u_st
                    st.session_state.students.at[idx, 'تاريخ التسجيل/Reg Date'] = u_rd
                    save_to_excel(st.session_state.students, st.session_state.ledger); st.success("Updated!")

# 4.3 الفوترة والتسويات
elif menu == T['menu'][2]:
    t1, t2, t3 = st.tabs([T['monthly_fees'], T['extra_hours'], T['adjustments']])
    with t1:
        m = st.selectbox("Month", range(1,13), index=datetime.date.today().month-1)
        y = st.number_input("Year", 2024, 2030, 2026)
        desc = f"Fees {m}-{y}"; last_d = datetime.date(y, m, calendar.monthrange(y, m)[1])
        if st.button(T['confirm'], key="bulk"):
            active = st.session_state.students[st.session_state.students['الحالة/Status']=='نشط']
            done = 0
            for _, s in active.iterrows():
                if s['تاريخ التسجيل/Reg Date'] <= last_d:
                    exists = not st.session_state.ledger[(st.session_state.ledger['اسم الطالب/Student'] == s['الاسم/Name']) & (st.session_state.ledger['البيان/Desc'] == desc)].empty
                    if not exists:
                        record_transaction(s['الاسم/Name'], "Invoice", desc, s['المبلغ الثابت/Fee'])
                        done += 1
            st.success(f"Billing Complete for {done} students!")
    with t2:
        with st.form("extra_form"):
            stb = st.selectbox(T['st_name'], st.session_state.students['الاسم/Name'] if not st.session_state.students.empty else [])
            col_a, col_b = st.columns(2)
            h = col_a.number_input("Hours", min_value=0.5, step=0.5)
            pr = col_b.number_input("Rate", min_value=0.0)
            ex_date = st.date_input("Service Date / تاريخ اليوم", datetime.date.today())
            note = st.text_input("Notes / ملاحظات")
            if st.form_submit_button(T['confirm']):
                desc_extra = f"ساعات إضافية ({h} س) بتاريخ {ex_date} - {note}"
                record_transaction(stb, "Invoice", desc_extra, h*pr, ex_date)
                st.success("Added Successfully!")
    with t3:
        with st.form("adj_form"):
            stb = st.selectbox(T['st_name'], st.session_state.students['الاسم/Name'] if not st.session_state.students.empty else [])
            col_x, col_y = st.columns(2)
            type_adj = col_x.selectbox("Type", ["Discount", "Debit"])
            amt = col_y.number_input("Amount", min_value=0.0)
            reason = st.text_input("Reason / السبب")
            if st.form_submit_button(T['confirm']):
                record_transaction(stb, type_adj, reason, amt)
                st.success("Done!")

# 4.4 التحصيل
elif menu == T['menu'][3]:
    if not st.session_state.students.empty:
        with st.form("pay_form"):
            stb = st.selectbox(T['st_name'], st.session_state.students['الاسم/Name'])
            amt = st.number_input("Amount", min_value=1.0)
            dt = st.date_input("Date", datetime.date.today())
            note = st.text_input("Notes")
            if st.form_submit_button(T['confirm']):
                record_transaction(stb, "Payment", note, amt, dt)
                st.success("Paid!")

# 4.5 التقارير
elif menu == T['menu'][4]:
    t_rep1, t_rep2 = st.tabs([T['summary_rep'], T['detailed_rep']])
    with t_rep1:
        st.dataframe(st.session_state.students, use_container_width=True)
        st.download_button(T['download'], to_excel(st.session_state.students), "Balances_Full.xlsx")
    with t_rep2:
        if not st.session_state.students.empty:
            stb = st.selectbox(T['st_name'], st.session_state.students['الاسم/Name'], key="audit_sel")
            res = st.session_state.ledger[st.session_state.ledger['اسم الطالب/Student']==stb]
            st.dataframe(res, use_container_width=True)
            if st.button(T['delete_last']):
                if not res.empty:
                    last_idx = res.index[-1]
                    s_idx = st.session_state.students[st.session_state.students['الاسم/Name'] == stb].index[0]
                    st.session_state.students.at[s_idx, 'الرصيد/Balance'] -= (st.session_state.ledger.at[last_idx, 'مدين/Debit (+)'] - st.session_state.ledger.at[last_idx, 'دائن/Credit (-)'])
                    st.session_state.ledger = st.session_state.ledger.drop(last_idx)
                    save_to_excel(st.session_state.students, st.session_state.ledger)
                    st.rerun()

# 4.6 تقارير الفترة
elif menu == T['menu'][5]:
    d1 = st.date_input("From", datetime.date.today()-datetime.timedelta(30))
    d2 = st.date_input("To", datetime.date.today())
    mask = (st.session_state.ledger['التاريخ/Date'].dt.date >= d1) & (st.session_state.ledger['التاريخ/Date'].dt.date <= d2)
    pdf = st.session_state.ledger[mask]
    st.dataframe(pdf, use_container_width=True)
    if not pdf.empty:
        st.download_button(T['download'], to_excel(pdf), "Period_Report.xlsx")