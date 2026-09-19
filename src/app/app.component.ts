import { CommonModule } from '@angular/common';
import { Component, OnInit } from '@angular/core';
import { FormsModule } from '@angular/forms';
import * as XLSX from 'xlsx';

/** shiftDate: تحديد يدوي لليوم الفعلي (الوردية) اللي تتبع له البصمة، بيتفعّل بس لو المستخدم نقلها يدويًا؛
 *  لو مش موجود بيتحسب تلقائيًا حسب قاعدة الوردية الليلية (انظر autoShiftDate). */
type Punch = { id: number; date: string; time: string; status: 'C/In' | 'C/Out'; location: string; shiftDate?: string };
/** توقيع اعتماد عام على بطاقة الموظف (اسم + تاريخ فقط، بدون تحديد مكان) */
type Signature = { id: number; name: string; date: string };
/** توقيع مرتبط بيوم معيّن: نوع الامضاء، تاريخه ووقته، اسم المكان، السبب، والملاحظات */
type DaySignature = { id: number; type: string; date: string; time: string; place: string; reason: string; notes: string };
type Employee = {
  no: string; name: string; department: string;
  punches: Punch[]; signatures: Signature[]; notes: string;
  dayNotes: Record<string, string>; daySignatures: Record<string, DaySignature[]>;
};
type ImportError = { row: number; message: string };
type DayStatus = 'complete' | 'incomplete' | 'absent';

@Component({
  selector: 'app-root', standalone: true, imports: [CommonModule, FormsModule],
  templateUrl: './app.component.html', styleUrl: './app.component.css'
})
export class AppComponent implements OnInit {
  employees: Employee[] = [];
  errors: ImportError[] = [];
  active?: Employee;
  view: 'dashboard' | 'employees' | 'employee' = 'dashboard';
  busy = false;
  savedFlash = false;
  private originals = new Map<number, Punch>();
  readonly requiredColumns = ['Department', 'Name', 'No', 'DateTime', 'Status', 'LocationID'];
  private readonly storageKey = 'mizan-state-v1';

  addDayOpen = false;
  signatureOpen = false;
  signatureName = '';

  daySigOpen = false;
  daySigTarget = '';
  readonly daySigTypes = ['توقيع حضور', 'توقيع انصراف', 'اعتماد مباشر', 'اعتماد إدارة', 'أخرى'];
  daySigForm = { type: '', date: '', time: '', place: '', reason: '', notes: '' };

  newDay = { date: '', inTime: '08:00', outTime: '16:00', inLocation: '', outLocation: '' };

  moveOpen = false;
  moveTarget?: Punch;
  moveDate = '';

  employeeSearch = '';
  private collapsedDepts = new Set<string>();

  ngOnInit() { this.restore(); }

  isCollapsed(dept: string) { return this.collapsedDepts.has(dept); }
  toggleDept(dept: string) { this.collapsedDepts.has(dept) ? this.collapsedDepts.delete(dept) : this.collapsedDepts.add(dept); }

  /** يقسّم الموظفين حسب القسم، ثم يفلتر بالبحث: لو البحث بيطابق اسم القسم يفضل القسم كامل،
   *  ولو بيطابق اسم/رقم موظف يفضل بس الموظفين المطابقين جوه أقسامهم. */
  get filteredGroups(): { department: string; employees: Employee[] }[] {
    const groupsMap = new Map<string, Employee[]>();
    this.employees.forEach(e => {
      const dept = e.department || 'بدون قسم';
      if (!groupsMap.has(dept)) groupsMap.set(dept, []);
      groupsMap.get(dept)!.push(e);
    });
    const groups = [...groupsMap.entries()]
      .map(([department, employees]) => ({ department, employees: [...employees].sort((a, b) => a.name.localeCompare(b.name, 'ar')) }))
      .sort((a, b) => a.department.localeCompare(b.department, 'ar'));

    const term = this.employeeSearch.trim().toLowerCase();
    if (!term) return groups;
    return groups
      .map(g => {
        const deptMatches = g.department.toLowerCase().includes(term);
        const employees = deptMatches ? g.employees : g.employees.filter(e => e.name.toLowerCase().includes(term) || e.no.toLowerCase().includes(term));
        return { department: g.department, employees };
      })
      .filter(g => g.employees.length > 0);
  }

  get totalHours() { return this.employees.reduce((n, e) => n + this.employeeHours(e), 0); }
  get totalDays() { return this.employees.reduce((n, e) => n + this.employeeDays(e), 0); }
  get totalAbsences() { return this.employees.reduce((n, e) => n + this.absentDays(e).length, 0); }
  get totalIssues() { return this.employees.reduce((n, e) => n + this.incompleteDays(e).length, 0); }

  employeeDays(e: Employee) { return [...new Set(e.punches.map(p => p.date))].filter(d => this.dayMinutes(e, d) !== null).length; }
  employeeHours(e: Employee) { return [...new Set(e.punches.map(p => p.date))].reduce((n, d) => n + (this.dayMinutes(e, d) || 0), 0) / 60; }
  days(e: Employee) { return [...new Set(e.punches.map(p => p.date))].sort(); }

  /** قاعدة الوردية الليلية لبصمة حضور: لو وقتها قبل 03:00 بتتحسب على يوم الوردية اللي قبلها (يوم أمس) */
  private autoShiftDate(date: string, time: string): string {
    if (time < '03:00') {
      const d = new Date(date + 'T00:00:00');
      d.setDate(d.getDate() - 1);
      return d.toISOString().slice(0, 10);
    }
    return date;
  }
  /** يربط كل بصمة انصراف بأقرب بصمة حضور فعلية قبلها زمنيًا (مش بس نفس التاريخ)،
   *  عشان انصراف الساعة 8 الصبح يقفل وردية أمس بالليل، مش يتحسب غلط مع حضور جديد الليلة دي.
   *  أي بصمة اتنقلت يدويًا (shiftDate) بتاخد يومها المحدد وبتخرج بره سلسلة الربط التلقائي. */
  private computeShiftDates(e: Employee): Map<number, string> {
    const map = new Map<number, string>();
    const sorted = [...e.punches].sort((a, b) => (a.date + a.time).localeCompare(b.date + b.time));
    let openShiftDate: string | null = null;
    for (const p of sorted) {
      if (p.shiftDate) { map.set(p.id, p.shiftDate); continue; }
      if (p.status === 'C/In') {
        openShiftDate = this.autoShiftDate(p.date, p.time);
        map.set(p.id, openShiftDate);
      } else if (openShiftDate) {
        map.set(p.id, openShiftDate); // بصمة انصراف بتقفل آخر وردية مفتوحة، أيًا كان تاريخها الفعلي
        openShiftDate = null;
      } else {
        map.set(p.id, this.autoShiftDate(p.date, p.time)); // انصراف بدون حضور سابق (بصمة يتيمة)
      }
    }
    return map;
  }
  /** اليوم الفعلي (يوم الوردية) اللي تتبع له البصمة: يدوي لو اتنقلت، وإلا حسب ربط الحضور بالانصراف */
  shiftDateOf(e: Employee, p: Punch): string { return this.computeShiftDates(e).get(p.id) ?? this.autoShiftDate(p.date, p.time); }
  /** هل البصمة دي منقولة يدويًا؟ */
  isMoved(p: Punch) { return !!p.shiftDate; }

  /** كل أيام الوردية من أول بصمة لآخر بصمة، بما فيها أيام الغياب، عشان يبان التاريخ كامل مش بس أيام الحضور */
  dateRange(e: Employee): string[] {
    const map = this.computeShiftDates(e);
    const dates = e.punches.map(p => map.get(p.id)!);
    if (!dates.length) return [];
    const sorted = [...dates].sort();
    const min = sorted[0], max = sorted[sorted.length - 1];
    const result: string[] = [];
    const cursor = new Date(min + 'T00:00:00');
    const end = new Date(max + 'T00:00:00');
    while (cursor <= end) {
      result.push(cursor.toISOString().slice(0, 10));
      cursor.setDate(cursor.getDate() + 1);
    }
    return result;
  }
  dayStatus(e: Employee, date: string): DayStatus {
    const map = this.computeShiftDates(e);
    const has = e.punches.some(p => map.get(p.id) === date);
    if (!has) return 'absent';
    return (this.firstIn(e, date) && this.lastOut(e, date)) ? 'complete' : 'incomplete';
  }
  /** نوع الوردية حسب توقيت الحضور الفعلي: من 18:00 لـ 06:00 تعتبر ليلية، وإلا صباحية */
  shiftType(e: Employee, date: string): 'ليلية' | 'صباحية' | null {
    const inPunch = this.firstIn(e, date);
    if (!inPunch) return null;
    const hour = +inPunch.time.split(':')[0];
    return (hour >= 18 || hour < 6) ? 'ليلية' : 'صباحية';
  }
  nightShiftDays(e: Employee) { return this.dateRange(e).filter(d => this.shiftType(e, d) === 'ليلية').length; }
  get totalNightShifts() { return this.employees.reduce((n, e) => n + this.nightShiftDays(e), 0); }
  statusLabel(s: DayStatus) { return s === 'complete' ? 'مكتمل' : s === 'incomplete' ? 'ناقص' : 'غياب'; }
  absentDays(e: Employee) { return this.dateRange(e).filter(d => this.dayStatus(e, d) === 'absent'); }
  incompleteDays(e: Employee) { return this.dateRange(e).filter(d => this.dayStatus(e, d) === 'incomplete'); }

  dayMinutes(e: Employee, date: string): number | null {
    const p = e.punches.filter(x => x.date === date).sort((a, b) => a.time.localeCompare(b.time));
    const first = p.find(x => x.status === 'C/In'); const last = [...p].reverse().find(x => x.status === 'C/Out');
    if (!first || !last) return null;
    let minutes = this.toMinutes(last.time) - this.toMinutes(first.time); if (minutes < 0) minutes += 1440;
    return minutes;
  }
  fmtMinutes(v: number | null) { if (v === null) return 'بصمة ناقصة'; return `${Math.floor(v / 60)} س ${v % 60} د`; }
  private toMinutes(time: string) { const [h, m] = time.split(':').map(Number); return h * 60 + m; }
  private fixText(value: unknown) {
    const text = String(value ?? '').trim();
    if (!/ط§|ظ„|ظ…/.test(text)) return text;
    try { return decodeURIComponent([...text].map(c => '%' + c.charCodeAt(0).toString(16).padStart(2, '0')).join('')); } catch { return text; }
  }

  importFile(event: Event) {
    const file = (event.target as HTMLInputElement).files?.[0]; if (!file) return;
    this.busy = true; this.errors = [];
    const reader = new FileReader();
    reader.onerror = () => { this.errors = [{ row: 0, message: 'تعذر قراءة الملف. تأكد من أنه ملف Excel أو CSV صالح.' }]; this.busy = false; };
    reader.onload = () => {
      try { this.readWorkbook(reader.result as ArrayBuffer); this.persist(); }
      catch (e) { this.errors = [{ row: 0, message: `تعذر استيراد الملف: ${e instanceof Error ? e.message : 'تنسيق غير مدعوم'}` }]; }
      finally { this.busy = false; }
    };
    reader.readAsArrayBuffer(file);
  }

  private readWorkbook(buffer: ArrayBuffer) {
    const wb = XLSX.read(buffer, { type: 'array', cellDates: true }); const sheet = wb.Sheets[wb.SheetNames[0]];
    if (!sheet) throw new Error('لا توجد ورقة بيانات في الملف.');
    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: '' });
    if (!rows.length) throw new Error('ورقة البيانات فارغة.');
    const headers = Object.keys(rows[0]); const missing = this.requiredColumns.filter(c => !headers.includes(c));
    if (missing.length) throw new Error(`الأعمدة المطلوبة غير موجودة: ${missing.join('، ')}.`);
    const map = new Map<string, Employee>(); let punchId = 1;
    rows.forEach((row, index) => {
      const no = String(row['No']).trim(), name = this.fixText(row['Name']), dept = this.fixText(row['Department']);
      const rawDate = row['DateTime']; const status = String(row['Status']).trim() as Punch['status'];
      const parsed = this.parseDateTime(rawDate);
      if (!no || !name || !parsed || !['C/In', 'C/Out'].includes(status)) { this.errors.push({ row: index + 2, message: 'بيانات ناقصة أو تاريخ/حالة غير صحيحة.' }); return; }
      const { date, time } = parsed;
      if (!map.has(no)) map.set(no, { no, name, department: dept, punches: [], signatures: [], notes: '', dayNotes: {}, daySignatures: {} });
      const punch = { id: punchId++, date, time, status, location: String(row['LocationID']).trim() };
      map.get(no)!.punches.push(punch); this.originals.set(punch.id, { ...punch });
    });
    this.employees = [...map.values()].sort((a, b) => a.name.localeCompare(b.name, 'ar'));
    if (!this.employees.length) throw new Error('لم يتم العثور على سجلات سليمة للاستيراد. راجع رسائل الأخطاء.');
    this.view = 'dashboard';
  }

  private parseDateTime(value: unknown): { date: string; time: string } | null {
    if (value instanceof Date && !isNaN(value.getTime())) return { date: `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, '0')}-${String(value.getDate()).padStart(2, '0')}`, time: `${String(value.getHours()).padStart(2, '0')}:${String(value.getMinutes()).padStart(2, '0')}` };
    const match = String(value).trim().match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{1,2}):(\d{2})/);
    if (!match || +match[4] > 23 || +match[5] > 59) return null;
    return { date: `${match[3]}-${match[2]}-${match[1]}`, time: `${match[4].padStart(2, '0')}:${match[5]}` };
  }

  select(e: Employee) { this.active = e; this.view = 'employee'; }

  openAddDay(prefillDate?: string) {
    this.newDay = { date: prefillDate || new Date().toISOString().slice(0, 10), inTime: '08:00', outTime: '16:00', inLocation: '', outLocation: '' };
    this.addDayOpen = true;
  }
  saveDay(e?: Employee) {
    if (!e || !this.newDay.date || !this.newDay.inTime || !this.newDay.outTime) return;
    const id = Date.now();
    // لو وقت الانصراف قبل وقت الحضور، معناها وردية ليلية عدّت نص الليل، فالانصراف فعليًا في اليوم التالي
    let outDate = this.newDay.date;
    if (this.newDay.outTime < this.newDay.inTime) {
      const d = new Date(this.newDay.date + 'T00:00:00'); d.setDate(d.getDate() + 1);
      outDate = d.toISOString().slice(0, 10);
    }
    e.punches.push(
      { id, date: this.newDay.date, time: this.newDay.inTime, status: 'C/In', location: this.newDay.inLocation },
      { id: id + 1, date: outDate, time: this.newDay.outTime, status: 'C/Out', location: this.newDay.outLocation }
    );
    this.addDayOpen = false;
    this.persist();
  }

  /** نقل بصمة معيّنة ليوم وردية تاني يدويًا، لو التقسيم التلقائي للوردية الليلية جاب لغبطة في حالة معيّنة */
  openMove(e: Employee, p: Punch) { this.moveTarget = p; this.moveDate = this.shiftDateOf(e, p); this.moveOpen = true; }
  saveMove() {
    if (!this.moveTarget || !this.moveDate) return;
    this.moveTarget.shiftDate = this.moveDate;
    this.moveOpen = false;
    this.persist();
  }
  resetMove(p: Punch) { delete p.shiftDate; this.persist(); }
  removePunch(e: Employee, p: Punch) { e.punches = e.punches.filter(x => x.id !== p.id); this.persist(); }
  dayPunches(e: Employee, date: string) {
    const map = this.computeShiftDates(e);
    return e.punches.filter(p => map.get(p.id) === date).sort((a, b) => (a.date + a.time).localeCompare(b.date + b.time));
  }
  firstIn(e: Employee, date: string) { return this.dayPunches(e, date).find(p => p.status === 'C/In'); }
  lastOut(e: Employee, date: string) { return [...this.dayPunches(e, date)].reverse().find(p => p.status === 'C/Out'); }
  isNew(p?: Punch) { return !!p && !this.originals.has(p.id); }
  isChanged(p: Punch) { const original = this.originals.get(p.id); return !!original && (original.date !== p.date || original.time !== p.time || original.status !== p.status || original.location !== p.location); }
  /** تبديل نوع البصمة الأصلية بين حضور وانصراف، لو جهاز البصمة سجّلها غلط */
  toggleStatus(p: Punch) { p.status = p.status === 'C/In' ? 'C/Out' : 'C/In'; this.persist(); }
  restorePunch(p: Punch) { const original = this.originals.get(p.id); if (original) Object.assign(p, original); this.persist(); }
  restoreDay(e: Employee, date: string) { this.dayPunches(e, date).forEach(p => this.restorePunch(p)); }
  roundPunch(p: Punch, direction: 'up' | 'down') {
    const [hours, minutes] = p.time.split(':').map(Number);
    let rounded = direction === 'up' ? (minutes ? hours + 1 : hours) : hours;
    rounded = ((rounded % 24) + 24) % 24; p.time = `${String(rounded).padStart(2, '0')}:00`;
    this.persist();
  }
  roundDay(e: Employee, date: string, direction: 'up' | 'down') { this.dayPunches(e, date).forEach(p => this.roundPunch(p, direction)); }

  /** توقيع اعتماد عام على البطاقة كلها: الاسم والتاريخ فقط، بدون تعديل على مكان — يظهر رقم ترتيبه للعلم فقط */
  openSignature() { this.signatureName = ''; this.signatureOpen = true; }
  saveSignature(e?: Employee) {
    if (!e || !this.signatureName.trim()) return;
    e.signatures.push({ id: Date.now(), name: this.signatureName.trim(), date: new Date().toISOString().slice(0, 10) });
    this.signatureOpen = false; this.signatureName = '';
    this.persist();
  }
  removeSignature(e: Employee, s: Signature) { e.signatures = e.signatures.filter(x => x.id !== s.id); this.persist(); }

  /** توقيع على مستوى يوم بعينه: نوع الامضاء + تاريخ ووقت + اسم المكان + السبب + ملاحظات */
  openDaySignature(date: string) {
    this.daySigTarget = date;
    this.daySigForm = { type: '', date, time: new Date().toTimeString().slice(0, 5), place: '', reason: '', notes: '' };
    this.daySigOpen = true;
  }
  saveDaySignature(e?: Employee) {
    if (!e || !this.daySigTarget || !this.daySigForm.type.trim()) return;
    if (!e.daySignatures[this.daySigTarget]) e.daySignatures[this.daySigTarget] = [];
    e.daySignatures[this.daySigTarget].push({ id: Date.now(), ...this.daySigForm });
    this.daySigOpen = false;
    this.persist();
  }
  removeDaySignature(e: Employee, date: string, ds: DaySignature) {
    e.daySignatures[date] = (e.daySignatures[date] || []).filter(x => x.id !== ds.id);
    this.persist();
  }
  /** ds كائن داخل المصفوفة وبيتعدل بالمرجع مباشرة عبر ngModel، الميثود دي بس تحفظ التعديل */
  updateDaySignature(_e: Employee, _date: string, _ds: DaySignature) { this.persist(); }
  daySignatures(e: Employee, date: string): DaySignature[] { return e.daySignatures[date] || []; }

  updateNote(_e: Employee, _value: string) { this.persist(); }
  updateDayNote(_e: Employee, _date: string, _value: string) { this.persist(); }

  persist() {
    try {
      const state = { employees: this.employees, originals: [...this.originals.entries()] };
      localStorage.setItem(this.storageKey, JSON.stringify(state));
    } catch { /* تجاهل أخطاء التخزين (مساحة ممتلئة أو متصفح يمنع الوصول) */ }
  }
  /** حفظ يدوي فوري مع رسالة تأكيد على الزرار، بالإضافة للحفظ التلقائي مع كل تعديل */
  saveNow() {
    this.persist();
    this.savedFlash = true;
    setTimeout(() => this.savedFlash = false, 2000);
  }
  private restore() {
    try {
      const raw = localStorage.getItem(this.storageKey);
      if (!raw) return;
      const state = JSON.parse(raw);
      this.employees = (state.employees || []).map((e: any) => ({ signatures: [], notes: '', dayNotes: {}, daySignatures: {}, ...e }));
      this.originals = new Map(state.originals || []);
    } catch { /* بيانات محفوظة تالفة، ابدأ فاضي */ }
  }
  clearSaved() {
    if (!confirm('هل تريد مسح كل البيانات المحفوظة من هذا الجهاز؟ لا يمكن التراجع عن هذا الإجراء.')) return;
    localStorage.removeItem(this.storageKey);
    this.employees = []; this.errors = []; this.active = undefined; this.view = 'dashboard'; this.originals.clear();
  }

  /** تصدير عام لكل بيانات الموظفين: نفس شكل ملف الاستيراد (Department, Name, No, DateTime, Status, LocationID)
   *  مع إضافة عمود البصمة الأصلية، وإظهار أيام الغياب في مكانها الصحيح بترتيب التاريخ، وكل موظف في مجموعة صفوف
   *  لوحده تنتهي بصف إجمالياته، مفصولة عن الموظف التالي بصف فارغ، وكل الموظفين تباعًا تحت بعض في نفس الشيت. */
  exportExcel() {
    const wb = XLSX.utils.book_new();

    // ---- ملخص سريع لكل موظف ----
    const summary = this.employees.map(e => ({
      'الرقم': e.no, 'الاسم': e.name, 'القسم': e.department,
      'أيام العمل': this.employeeDays(e), 'أيام الغياب': this.absentDays(e).length, 'أيام ناقصة': this.incompleteDays(e).length,
      'ساعات العمل': +this.employeeHours(e).toFixed(2),
      'أيام وردية ليلية': this.nightShiftDays(e),
      'عدد التوقيعات العامة': e.signatures.length,
      'ملاحظات': e.notes || '—'
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), 'الرواتب');

    // ---- سحبة البصمة الكاملة: نفس شكل الاستيراد + عمود الأصلي ونوع الوردية، وأيام الغياب في ترتيبها ----
    const header = ['Department', 'Name', 'No', 'DateTime', 'Status', 'LocationID', 'DateTime الأصلي', 'نوع الوردية', 'ملاحظة'];
    const rows: (string | number)[][] = [header];
    this.employees.forEach(e => {
      this.dateRange(e).forEach(date => {
        const punches = this.dayPunches(e, date);
        if (!punches.length) {
          rows.push([e.department, e.name, e.no, date.split('-').reverse().join('/'), 'غياب', '', '', '', e.dayNotes[date] || '']);
          return;
        }
        const shift = this.shiftType(e, date) || '';
        punches.forEach(p => {
          const original = this.originals.get(p.id);
          const current = `${p.date.split('-').reverse().join('/')} ${p.time}`;
          const orig = original ? `${original.date.split('-').reverse().join('/')} ${original.time}` : current;
          rows.push([e.department, e.name, e.no, current, p.status, p.location, orig, shift, e.dayNotes[date] || '']);
        });
      });
      rows.push(['', `الإجمالي — أيام العمل: ${this.employeeDays(e)} | أيام غياب: ${this.absentDays(e).length} | أيام ناقصة: ${this.incompleteDays(e).length} | ورديات ليلية: ${this.nightShiftDays(e)} | الساعات: ${this.employeeHours(e).toFixed(1)} س`, '', '', '', '', '', '', e.notes || '']);
      rows.push([]); // صف فارغ يفصل بين مربع كل موظف والتالي
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'سحبة البصمة الكاملة');

    // ---- كل التوقيعات: العامة وتوقيعات الأيام ----
    const sigRows: (string | number)[][] = [['الرقم', 'الاسم', 'النوع', 'التاريخ', 'الوقت', 'المكان', 'السبب', 'ملاحظات']];
    this.employees.forEach(e => {
      e.signatures.forEach((s, i) => sigRows.push([e.no, e.name, `توقيع اعتماد عام رقم ${i + 1}`, s.date, '', '', '', '']));
      Object.keys(e.daySignatures).sort().forEach(date => {
        (e.daySignatures[date] || []).forEach(ds => sigRows.push([e.no, e.name, ds.type, ds.date, ds.time, ds.place, ds.reason, ds.notes]));
      });
    });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sigRows), 'التوقيعات');

    // ---- أخطاء الاستيراد ----
    const errRows: (string | number)[][] = [['الصف', 'الخطأ'], ...this.errors.map(x => [x.row, x.message])];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(errRows), 'أخطاء الاستيراد');

    XLSX.writeFile(wb, 'تقرير_الحضور_والرواتب.xlsx');
  }

  print() { window.print(); }
  trackDate(_i: number, d: string) { return d; }
  trackPunch(_i: number, p: Punch) { return p.id; }
  trackSig(_i: number, s: Signature) { return s.id; }
  trackDaySig(_i: number, s: DaySignature) { return s.id; }
}