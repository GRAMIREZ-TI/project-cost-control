"""Project cost control: baseline, commitments, actuals, changes and earned value."""
import os, io, decimal
from datetime import date, datetime
from contextlib import asynccontextmanager
from fastapi import FastAPI, Depends, HTTPException, Form, Request
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse
from fastapi.templating import Jinja2Templates
from sqlalchemy import create_engine, Column, Integer, String, Date, DateTime, Numeric, ForeignKey, UniqueConstraint, CheckConstraint, func, select
from sqlalchemy.orm import declarative_base, sessionmaker, Session, relationship
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape

url = os.getenv('DATABASE_URL', 'sqlite:///./project_costs.db')
if url.startswith('postgres://'): url = 'postgresql+psycopg://' + url[len('postgres://'):]
if url.startswith('postgresql://'): url = 'postgresql+psycopg://' + url[len('postgresql://'):]
engine = create_engine(url, connect_args={'check_same_thread': False} if url.startswith('sqlite') else {}, pool_pre_ping=True)
SessionLocal = sessionmaker(bind=engine)
Base = declarative_base()
D = decimal.Decimal

class Project(Base):
    __tablename__ = 'projects'
    id = Column(Integer, primary_key=True)
    code = Column(String(40), unique=True, nullable=False)
    name = Column(String(200), nullable=False)
    currency = Column(String(3), nullable=False, default='PEN')
    status = Column(String(20), nullable=False, default='ACTIVE')

class WorkPackage(Base):
    __tablename__ = 'work_packages'
    id = Column(Integer, primary_key=True)
    project_id = Column(Integer, ForeignKey('projects.id'), nullable=False)
    code = Column(String(40), nullable=False)
    name = Column(String(200), nullable=False)
    __table_args__ = (UniqueConstraint('project_id','code'),)

class Baseline(Base):
    __tablename__ = 'baselines'
    id = Column(Integer, primary_key=True)
    project_id = Column(Integer, ForeignKey('projects.id'), nullable=False)
    version = Column(Integer, nullable=False)
    status = Column(String(20), nullable=False, default='DRAFT')
    approved_at = Column(DateTime)
    __table_args__ = (UniqueConstraint('project_id','version'),)

class BudgetLine(Base):
    __tablename__ = 'budget_lines'
    id = Column(Integer, primary_key=True)
    baseline_id = Column(Integer, ForeignKey('baselines.id'), nullable=False)
    package_id = Column(Integer, ForeignKey('work_packages.id'), nullable=False)
    period = Column(Date, nullable=False)
    amount = Column(Numeric(18,2), nullable=False)
    __table_args__ = (UniqueConstraint('baseline_id','package_id','period'), CheckConstraint('amount >= 0'),)

class Commitment(Base):
    __tablename__ = 'commitments'
    id = Column(Integer, primary_key=True)
    project_id = Column(Integer, ForeignKey('projects.id'), nullable=False)
    package_id = Column(Integer, ForeignKey('work_packages.id'), nullable=False)
    reference = Column(String(80), nullable=False)
    supplier = Column(String(150), nullable=False)
    amount = Column(Numeric(18,2), nullable=False)
    committed_on = Column(Date, nullable=False)
    __table_args__ = (UniqueConstraint('project_id','reference'), CheckConstraint('amount >= 0'),)

class Actual(Base):
    __tablename__ = 'actuals'
    id = Column(Integer, primary_key=True)
    project_id = Column(Integer, ForeignKey('projects.id'), nullable=False)
    package_id = Column(Integer, ForeignKey('work_packages.id'), nullable=False)
    commitment_id = Column(Integer, ForeignKey('commitments.id'))
    reference = Column(String(80), nullable=False)
    amount = Column(Numeric(18,2), nullable=False)
    incurred_on = Column(Date, nullable=False)
    __table_args__ = (UniqueConstraint('project_id','reference'), CheckConstraint('amount >= 0'),)

class Change(Base):
    __tablename__ = 'changes'
    id = Column(Integer, primary_key=True)
    project_id = Column(Integer, ForeignKey('projects.id'), nullable=False)
    package_id = Column(Integer, ForeignKey('work_packages.id'), nullable=False)
    reference = Column(String(80), nullable=False)
    amount_delta = Column(Numeric(18,2), nullable=False)
    status = Column(String(20), nullable=False, default='PENDING')
    approved_on = Column(Date)
    __table_args__ = (UniqueConstraint('project_id','reference'),)

class Progress(Base):
    __tablename__ = 'progress'
    id = Column(Integer, primary_key=True)
    package_id = Column(Integer, ForeignKey('work_packages.id'), nullable=False)
    measured_on = Column(Date, nullable=False)
    percent_complete = Column(Numeric(5,2), nullable=False)
    __table_args__ = (UniqueConstraint('package_id','measured_on'), CheckConstraint('percent_complete >= 0 AND percent_complete <= 100'),)

@asynccontextmanager
async def lifespan(app):
    Base.metadata.create_all(engine)
    yield

app = FastAPI(title='Capacidad 2 | Project Cost Control', lifespan=lifespan)
templates = Jinja2Templates(directory='templates')

def db_session():
    with SessionLocal() as db: yield db

def project_or_404(db, project_id):
    p = db.get(Project, project_id)
    if not p: raise HTTPException(404, 'Proyecto no encontrado')
    return p

def package_or_404(db, project_id, package_id):
    p = db.get(WorkPackage, package_id)
    if not p or p.project_id != project_id: raise HTTPException(400, 'Paquete WBS no pertenece al proyecto')
    return p

def money(x): return D(str(x or 0))
def total(db, model, column, *filters):
    return money(db.scalar(select(func.sum(column)).where(*filters)))

def metrics(db, project_id, cutoff=None):
    project_or_404(db, project_id)
    cutoff = cutoff or date.today()
    approved = db.scalars(select(Baseline).where(Baseline.project_id==project_id, Baseline.status=='APPROVED').order_by(Baseline.version.desc())).first()
    packages = db.scalars(select(WorkPackage).where(WorkPackage.project_id==project_id).order_by(WorkPackage.code)).all()
    rows=[]
    for w in packages:
        bac = total(db, BudgetLine, BudgetLine.amount, BudgetLine.baseline_id==approved.id, BudgetLine.package_id==w.id) if approved else D('0')
        pv = total(db, BudgetLine, BudgetLine.amount, BudgetLine.baseline_id==approved.id, BudgetLine.package_id==w.id, BudgetLine.period<=cutoff) if approved else D('0')
        commitment = total(db, Commitment, Commitment.amount, Commitment.package_id==w.id, Commitment.committed_on<=cutoff)
        ac = total(db, Actual, Actual.amount, Actual.package_id==w.id, Actual.incurred_on<=cutoff)
        changes = total(db, Change, Change.amount_delta, Change.package_id==w.id, Change.status=='APPROVED', Change.approved_on<=cutoff)
        progress = db.scalars(select(Progress).where(Progress.package_id==w.id, Progress.measured_on<=cutoff).order_by(Progress.measured_on.desc(), Progress.id.desc())).first()
        pct = money(progress.percent_complete) if progress else D('0')
        ev = bac*pct/D('100')
        rows.append(dict(code=w.code, name=w.name, bac=bac, pv=pv, ev=ev, ac=ac, commitments=commitment, changes=changes, pct=pct, cv=ev-ac, sv=ev-pv, revised=bac+changes, exposure=commitment-ac))
    sums={k:sum((r[k] for r in rows), D('0')) for k in ('bac','pv','ev','ac','commitments','changes','cv','sv','revised','exposure')}
    sums['cpi'] = sums['ev']/sums['ac'] if sums['ac'] else None
    sums['spi'] = sums['ev']/sums['pv'] if sums['pv'] else None
    sums['eac'] = sums['ac']+(sums['bac']-sums['ev'])/sums['cpi'] if sums['cpi'] and sums['cpi']>0 else None
    sums['vac'] = sums['bac']-sums['eac'] if sums['eac'] is not None else None
    return approved, rows, sums

def redirect(pid): return RedirectResponse(f'/projects/{pid}', status_code=303)

def render(request, db, pid=None, error=None):
    projects=db.scalars(select(Project).order_by(Project.code)).all()
    p=project_or_404(db,pid) if pid else None
    packages=db.scalars(select(WorkPackage).where(WorkPackage.project_id==pid).order_by(WorkPackage.code)).all() if p else []
    baselines=db.scalars(select(Baseline).where(Baseline.project_id==pid).order_by(Baseline.version.desc())).all() if p else []
    approved,rows,sums=metrics(db,pid) if p else (None,[],{})
    recent={}
    if p:
        for name,model in [('commitments',Commitment),('actuals',Actual),('changes',Change)]:
            recent[name]=db.scalars(select(model).where(model.project_id==pid).order_by(model.id.desc()).limit(10)).all()
    return templates.TemplateResponse(request, 'index.html', dict(projects=projects, project=p, packages=packages, baselines=baselines, approved=approved, rows=rows, sums=sums, recent=recent, error=error, today=date.today().isoformat()))

@app.get('/', response_class=HTMLResponse)
def home(request:Request, db:Session=Depends(db_session)): return render(request,db)
@app.get('/projects/{pid}', response_class=HTMLResponse)
def detail(pid:int,request:Request,db:Session=Depends(db_session)): return render(request,db,pid)

@app.post('/projects')
def add_project(code:str=Form(...), name:str=Form(...), currency:str=Form('PEN'), db:Session=Depends(db_session)):
    if db.scalar(select(Project).where(Project.code==code.strip())): raise HTTPException(409,'Código duplicado')
    p=Project(code=code.strip(),name=name.strip(),currency=currency.upper());db.add(p);db.commit();db.refresh(p)
    return redirect(p.id)

@app.post('/projects/{pid}/packages')
def add_package(pid:int,code:str=Form(...),name:str=Form(...),db:Session=Depends(db_session)):
    project_or_404(db,pid)
    if db.scalar(select(WorkPackage).where(WorkPackage.project_id==pid,WorkPackage.code==code.strip())): raise HTTPException(409,'WBS duplicada')
    db.add(WorkPackage(project_id=pid,code=code.strip(),name=name.strip()));db.commit();return redirect(pid)

@app.post('/projects/{pid}/baselines')
def add_baseline(pid:int,db:Session=Depends(db_session)):
    project_or_404(db,pid)
    version=(db.scalar(select(func.max(Baseline.version)).where(Baseline.project_id==pid)) or 0)+1
    db.add(Baseline(project_id=pid,version=version));db.commit();return redirect(pid)

@app.post('/projects/{pid}/budgets')
def add_budget(pid:int,baseline_id:int=Form(...),package_id:int=Form(...),period:date=Form(...),amount:decimal.Decimal=Form(...),db:Session=Depends(db_session)):
    b=db.get(Baseline,baseline_id);package_or_404(db,pid,package_id)
    if not b or b.project_id!=pid or b.status!='DRAFT':raise HTTPException(400,'Solo se modifica una línea base en borrador')
    if amount<0:raise HTTPException(400,'Monto negativo')
    existing=db.scalar(select(BudgetLine).where(BudgetLine.baseline_id==baseline_id,BudgetLine.package_id==package_id,BudgetLine.period==period))
    if existing:existing.amount=amount
    else:db.add(BudgetLine(baseline_id=baseline_id,package_id=package_id,period=period,amount=amount))
    db.commit();return redirect(pid)

@app.post('/projects/{pid}/baselines/{bid}/approve')
def approve_baseline(pid:int,bid:int,db:Session=Depends(db_session)):
    b=db.get(Baseline,bid)
    if not b or b.project_id!=pid or b.status!='DRAFT':raise HTTPException(400,'Línea base inválida')
    if not db.scalar(select(BudgetLine.id).where(BudgetLine.baseline_id==bid).limit(1)):raise HTTPException(400,'La línea base no tiene partidas')
    for prev in db.scalars(select(Baseline).where(Baseline.project_id==pid,Baseline.status=='APPROVED')):prev.status='SUPERSEDED'
    b.status='APPROVED';b.approved_at=datetime.utcnow();db.commit();return redirect(pid)

@app.post('/projects/{pid}/commitments')
def add_commitment(pid:int,package_id:int=Form(...),reference:str=Form(...),supplier:str=Form(...),amount:decimal.Decimal=Form(...),committed_on:date=Form(...),db:Session=Depends(db_session)):
    package_or_404(db,pid,package_id)
    if amount<0:raise HTTPException(400,'Monto negativo')
    if db.scalar(select(Commitment.id).where(Commitment.project_id==pid,Commitment.reference==reference)):raise HTTPException(409,'Referencia duplicada')
    db.add(Commitment(project_id=pid,package_id=package_id,reference=reference,supplier=supplier,amount=amount,committed_on=committed_on));db.commit();return redirect(pid)

@app.post('/projects/{pid}/actuals')
def add_actual(pid:int,package_id:int=Form(...),reference:str=Form(...),amount:decimal.Decimal=Form(...),incurred_on:date=Form(...),commitment_id:int|None=Form(None),db:Session=Depends(db_session)):
    package_or_404(db,pid,package_id)
    if amount<0:raise HTTPException(400,'Monto negativo')
    if db.scalar(select(Actual.id).where(Actual.project_id==pid,Actual.reference==reference)):raise HTTPException(409,'Referencia duplicada')
    if commitment_id:
        c=db.get(Commitment,commitment_id)
        if not c or c.project_id!=pid or c.package_id!=package_id:raise HTTPException(400,'Compromiso incompatible')
    db.add(Actual(project_id=pid,package_id=package_id,reference=reference,amount=amount,incurred_on=incurred_on,commitment_id=commitment_id));db.commit();return redirect(pid)

@app.post('/projects/{pid}/changes')
def add_change(pid:int,package_id:int=Form(...),reference:str=Form(...),amount_delta:decimal.Decimal=Form(...),db:Session=Depends(db_session)):
    package_or_404(db,pid,package_id)
    if db.scalar(select(Change.id).where(Change.project_id==pid,Change.reference==reference)):raise HTTPException(409,'Referencia duplicada')
    db.add(Change(project_id=pid,package_id=package_id,reference=reference,amount_delta=amount_delta));db.commit();return redirect(pid)

@app.post('/projects/{pid}/changes/{cid}/approve')
def approve_change(pid:int,cid:int,db:Session=Depends(db_session)):
    c=db.get(Change,cid)
    if not c or c.project_id!=pid or c.status!='PENDING':raise HTTPException(400,'Cambio inválido')
    c.status='APPROVED';c.approved_on=date.today();db.commit();return redirect(pid)

@app.post('/projects/{pid}/progress')
def add_progress(pid:int,package_id:int=Form(...),measured_on:date=Form(...),percent_complete:decimal.Decimal=Form(...),db:Session=Depends(db_session)):
    package_or_404(db,pid,package_id)
    if not D('0')<=percent_complete<=D('100'):raise HTTPException(400,'Avance fuera de rango')
    progress=db.scalar(select(Progress).where(Progress.package_id==package_id,Progress.measured_on==measured_on))
    if progress:progress.percent_complete=percent_complete
    else:db.add(Progress(package_id=package_id,measured_on=measured_on,percent_complete=percent_complete))
    db.commit();return redirect(pid)

HEADERS=['WBS','Paquete','BAC','PV','EV','AC','Compromisos','Cambios aprobados','Presupuesto ajustado','Avance %','CV','SV']
def report_data(rows):
    return [[r['code'],r['name']]+[float(r[k]) for k in ('bac','pv','ev','ac','commitments','changes','revised','pct','cv','sv')] for r in rows]

@app.get('/projects/{pid}/export.xlsx')
def export_excel(pid:int,db:Session=Depends(db_session)):
    p=project_or_404(db,pid);_,rows,s=metrics(db,pid)
    wb=Workbook();ws=wb.active;ws.title='Control de costos'
    ws.merge_cells('A1:L1');ws['A1']='CONTROL PRESUPUESTAL Y VALOR GANADO | '+p.name
    ws['A1'].font=Font(size=16,bold=True,color='FFFFFF');ws['A1'].fill=PatternFill('solid',fgColor='102B4E');ws.row_dimensions[1].height=32
    ws['A2']=f'Proyecto: {p.code} | Moneda: {p.currency} | Corte: {date.today():%d/%m/%Y}'
    ws.append([]);ws.append(HEADERS)
    for c in ws[4]:c.font=Font(color='FFFFFF',bold=True);c.fill=PatternFill('solid',fgColor='176B9B')
    for record in report_data(rows):ws.append(record)
    end=ws.max_row+1;ws.cell(end,2,'TOTAL')
    for idx,key in [(3,'bac'),(4,'pv'),(5,'ev'),(6,'ac'),(7,'commitments'),(8,'changes'),(9,'revised'),(11,'cv'),(12,'sv')]:ws.cell(end,idx,float(s[key]))
    for c in ws[end]:c.font=Font(bold=True,color='102B4E');c.fill=PatternFill('solid',fgColor='DCEAF3')
    ws.append([]);ws.append(['CPI',float(s['cpi']) if s['cpi'] is not None else 'N/D','SPI',float(s['spi']) if s['spi'] is not None else 'N/D','EAC',float(s['eac']) if s['eac'] is not None else 'N/D','VAC',float(s['vac']) if s['vac'] is not None else 'N/D'])
    ws.append(['Nota: BAC/PV/EV se calculan sobre la línea base aprobada; cambios aprobados se muestran aparte.'])
    for col in range(1,13):ws.column_dimensions[get_column_letter(col)].width=21 if col>2 else (16 if col==1 else 36)
    for row in ws.iter_rows(min_row=5,max_row=end,min_col=3,max_col=12):
        for cell in row:cell.number_format='#,##0.00;[Red](#,##0.00)'
    ws.freeze_panes='C5';ws.auto_filter.ref=f'A4:L{end}';ws.sheet_view.showGridLines=False
    out=io.BytesIO();wb.save(out);out.seek(0)
    return StreamingResponse(out,media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={'Content-Disposition':f'attachment; filename="costos_{p.code}.xlsx"'})

@app.get('/projects/{pid}/export.pdf')
def export_pdf(pid:int,db:Session=Depends(db_session)):
    p=project_or_404(db,pid);_,rows,s=metrics(db,pid)
    out=io.BytesIO();doc=SimpleDocTemplate(out,pagesize=landscape(A4),leftMargin=26,rightMargin=26,topMargin=28,bottomMargin=28)
    styles=getSampleStyleSheet();story=[Paragraph('CONTROL PRESUPUESTAL Y VALOR GANADO',styles['Title']),Paragraph(f'{p.code} | {p.name} | {p.currency} | Corte {date.today():%d/%m/%Y}',styles['Normal']),Spacer(1,14)]
    cols=['WBS','Paquete','BAC','PV','EV','AC','Comp.','Cambios','CV','SV']
    keys=['code','name','bac','pv','ev','ac','commitments','changes','cv','sv']
    data=[cols]
    for r in rows:data.append([str(r[k]) if k in ('code','name') else f'{r[k]:,.2f}' for k in keys])
    data.append(['','TOTAL']+[f'{s[k]:,.2f}' for k in keys[2:]])
    table=Table(data,repeatRows=1,colWidths=[48,165]+[66]*8,hAlign='LEFT')
    table.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#102B4E')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('BACKGROUND',(0,-1),(-1,-1),colors.HexColor('#DCEAF3')),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),7),('BOTTOMPADDING',(0,0),(-1,0),9),('GRID',(0,0),(-1,-1),0.25,colors.HexColor('#D3DDE7'))]))
    story.append(table);story.append(Spacer(1,14))
    fmt=lambda v:f'{v:,.2f}' if v is not None else 'N/D'
    story.append(Paragraph(f"CPI: {fmt(s['cpi'])} &nbsp;&nbsp; SPI: {fmt(s['spi'])} &nbsp;&nbsp; EAC: {fmt(s['eac'])} &nbsp;&nbsp; VAC: {fmt(s['vac'])}",styles['Heading3']))
    story.append(Paragraph('Metodología: BAC/PV/EV provienen de la línea base aprobada; los cambios aprobados no modifican retrospectivamente la línea base. Los compromisos no se suman al costo real.',styles['Normal']))
    doc.build(story);out.seek(0)
    return StreamingResponse(out,media_type='application/pdf',headers={'Content-Disposition':f'attachment; filename="costos_{p.code}.pdf"'})

@app.get('/health')
def health():return {'status':'ok'}
