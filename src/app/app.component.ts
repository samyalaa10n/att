import { CommonModule } from '@angular/common';
import { Component } from '@angular/core';
import { FormsModule } from '@angular/forms';
import * as XLSX from 'xlsx';

type Punch = { id: number; date: string; time: string; status: 'C/In' | 'C/Out'; location: string };
type Employee = { no: string; name: string; department: string; punches: Punch[]; basic: number; allowance: number };
type ImportError = { row: number; message: string };

@Component({
  selector: 'app-root', standalone: true, imports: [CommonModule, FormsModule],
  templateUrl: './app.component.html', styleUrl: './app.component.css'
})
export class AppComponent {
  employees: Employee[] = [];
  errors: ImportError[] = [];
  active?: Employee;
  view: 'dashboard' | 'employees' | 'employee' = 'dashboard';
  busy = false;
  readonly requiredColumns = ['Department', 'Name', 'No', 'DateTime', 'Status', 'LocationID'];

  get totalHours() { return this.employees.reduce((n, e) => n + this.employeeHours(e), 0); }
  get totalDays() { return this.employees.reduce((n, e) => n + this.employeeDays(e), 0); }
  get totalPayroll() { return this.employees.reduce((n, e) => n + (+e.basic || 0) + (+e.allowance || 0), 0); }
  employeeDays(e: Employee) { return [...new Set(e.punches.map(p => p.date))].filter(d => this.dayMinutes(e, d) !== null).length; }
  employeeHours(e: Employee) { return [...new Set(e.punches.map(p => p.date))].reduce((n, d) => n + (this.dayMinutes(e, d) || 0), 0) / 60; }
  days(e: Employee) { return [...new Set(e.punches.map(p => p.date))].sort(); }
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
    reader.onload = () => { try { this.readWorkbook(reader.result as ArrayBuffer); } catch (e) { this.errors = [{ row: 0, message: `تعذر استيراد الملف: ${e instanceof Error ? e.message : 'تنسيق غير مدعوم'}` }]; } finally { this.busy = false; } };
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
      if (!map.has(no)) map.set(no, { no, name, department: dept, punches: [], basic: 0, allowance: 0 });
      map.get(no)!.punches.push({ id: punchId++, date, time, status, location: String(row['LocationID']).trim() });
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
  addPunch(e: Employee) { e.punches.push({ id: Date.now(), date: new Date().toISOString().slice(0, 10), time: '08:00', status: 'C/In', location: '' }); }
  removePunch(e: Employee, p: Punch) { e.punches = e.punches.filter(x => x.id !== p.id); }
  exportExcel() {
    const wb = XLSX.utils.book_new();
    const employees = this.employees.map(e => ({ 'الرقم': e.no, 'الاسم': e.name, 'القسم': e.department, 'أيام العمل': this.employeeDays(e), 'ساعات العمل': +this.employeeHours(e).toFixed(2), 'الأساسي': e.basic, 'البدلات': e.allowance, 'إجمالي الراتب': (+e.basic || 0) + (+e.allowance || 0) }));
    const attendance = this.employees.flatMap(e => e.punches.map(p => ({ 'الرقم': e.no, 'الاسم': e.name, 'التاريخ': p.date, 'الوقت': p.time, 'الحالة': p.status, 'الموقع': p.location })));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(employees), 'الرواتب');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(attendance), 'الحضور والبصمات');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(this.errors.map(x => ({ 'الصف': x.row, 'الخطأ': x.message }))), 'أخطاء الاستيراد');
    XLSX.writeFile(wb, 'تقرير_الحضور_والرواتب.xlsx');
  }
  print() { window.print(); }
}
