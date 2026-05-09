require('dotenv').config();
const express = require('express');
const cors = require('cors');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const jwt = require('jsonwebtoken');
const bcrypt = require('bcryptjs');
const cron = require('node-cron');
const { v4: uuidv4 } = require('uuid');
const nodemailer = require('nodemailer');
const ExcelJS = require('exceljs');
const { initDb, saveToDisk } = require('./database');

const app = express();
const PORT = process.env.PORT || 3000;
const JWT_SECRET = process.env.JWT_SECRET || 'flp-secreto-cambia-esto-2024';
const UPLOADS_DIR = path.join(__dirname, 'uploads');
const BACKUP_DIR = path.join(__dirname, 'backups');
[UPLOADS_DIR, BACKUP_DIR].forEach(d => { if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true }); });

app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));
app.use('/uploads', express.static(UPLOADS_DIR));

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOADS_DIR),
  filename: (req, file, cb) => cb(null, uuidv4() + path.extname(file.originalname))
});
const upload = multer({ storage, limits: { fileSize: 5 * 1024 * 1024 } });

// ─── Utils ────────────────────────────────────────────────────
function authMw(req, res, next) {
  const token = req.headers.authorization?.split(' ')[1];
  if (!token) return res.status(401).json({ error: 'Token requerido' });
  try { req.user = jwt.verify(token, JWT_SECRET); next(); }
  catch { res.status(401).json({ error: 'Token inválido o expirado' }); }
}
function soloSuperAdmin(req, res, next) {
  if (req.user.rol !== 'admin' || !req.user.es_superadmin)
    return res.status(403).json({ error: 'Solo el administrador principal puede realizar esta acción' });
  next();
}
function soloAdmin(req, res, next) {
  if (req.user.rol !== 'admin') return res.status(403).json({ error: 'Solo administradores' });
  next();
}

// Validar RUT chileno
function validarRut(numero, dv) {
  const rut = parseInt(numero.replace(/\D/g,''));
  if (isNaN(rut) || rut < 1000000) return false;
  let suma = 0, multiplo = 2;
  let tmp = rut;
  while (tmp > 0) {
    suma += (tmp % 10) * multiplo;
    tmp = Math.floor(tmp / 10);
    multiplo = multiplo < 7 ? multiplo + 1 : 2;
  }
  const dvEsperado = 11 - (suma % 11);
  const dvCalc = dvEsperado === 11 ? '0' : dvEsperado === 10 ? 'K' : String(dvEsperado);
  return dvCalc === dv.toUpperCase();
}

function formatRut(numero, dv) {
  const n = numero.replace(/\D/g,'');
  return n.replace(/\B(?=(\d{3})+(?!\d))/g, '.') + '-' + dv.toUpperCase();
}

function calcEstado(hora, esperada) {
  if (!hora) return { estado: 'ausente', minutos: 0 };
  const [eh,em] = esperada.split(':').map(Number);
  const [ah,am] = hora.split(':').map(Number);
  const diff = (ah*60+am) - (eh*60+em);
  if (diff <= 5) return { estado: 'presente', minutos: 0 };
  if (diff <= 30) return { estado: 'tardanza_leve', minutos: diff };
  return { estado: 'tardanza_grave', minutos: diff };
}

function calcHorasMinutos(entrada, salida, horasSemanales = 45) {
  if (!entrada || !salida) return { ordinarias: 0, extra: 0, total: 0 };
  const [eh,em] = entrada.split(':').map(Number);
  const [sh,sm] = salida.split(':').map(Number);
  const totalMins = (sh*60+sm) - (eh*60+em);
  if (totalMins <= 0) return { ordinarias: 0, extra: 0, total: 0 };
  const totalH = totalMins / 60;
  // Horas diarias ordinarias máx según contrato (45h/semana = 9h/día en 5 días)
  const maxDiarias = (horasSemanales || 45) / 5;
  const ordinarias = Math.min(totalH, maxDiarias);
  const extra = Math.max(0, totalH - maxDiarias);
  return { ordinarias: Math.round(ordinarias*100)/100, extra: Math.round(extra*100)/100, total: Math.round(totalH*100)/100 };
}

function guardarFoto(b64) {
  if (!b64 || !b64.startsWith('data:image')) return null;
  const m = b64.match(/^data:image\/(\w+);base64,(.+)$/); if (!m) return null;
  const fname = uuidv4()+'.'+m[1];
  fs.writeFileSync(path.join(UPLOADS_DIR, fname), Buffer.from(m[2],'base64'));
  return fname;
}

function log(db, usuario, rol, accion, detalle, ip) {
  try { db.prepare("INSERT INTO logs (usuario,rol,accion,detalle,ip) VALUES (?,?,?,?,?)").run(usuario||'sistema',rol||'sistema',accion,detalle||'',ip||''); }
  catch(e) {}
}

// ─── Respaldo automático ──────────────────────────────────────
function hacerRespaldo(db) {
  try {
    const fecha = new Date().toISOString().slice(0,10);
    const archivo = path.join(BACKUP_DIR, `asistencia_backup_${fecha}.db`);
    saveToDisk();
    const DB_PATH = process.env.DB_PATH || path.join(__dirname, 'asistencia.db');
    if (fs.existsSync(DB_PATH)) {
      fs.copyFileSync(DB_PATH, archivo);
      const stats = fs.statSync(archivo);
      db.prepare('INSERT INTO respaldos (archivo,tamanio) VALUES (?,?)').run(archivo, stats.size);
      // Conservar solo últimos 30 respaldos
      const respaldos = db.prepare('SELECT * FROM respaldos ORDER BY creado_en DESC').all();
      if (respaldos.length > 30) {
        const viejos = respaldos.slice(30);
        viejos.forEach(r => { try { fs.unlinkSync(r.archivo); db.prepare('DELETE FROM respaldos WHERE id=?').run(r.id); } catch(e){} });
      }
      console.log(`✅ Respaldo creado: ${archivo}`);
    }
  } catch(e) { console.error('Error en respaldo:', e.message); }
}

// ─── Excel libro de asistencia ────────────────────────────────
async function generarLibroMensual(db, mes, obraId) {
  const [anio, m] = mes.split('-');
  const registros = db.prepare(`
    SELECT r.*, e.nombre, e.rut_numero, e.rut_dv, e.cargo, e.hora_entrada as hora_esp, e.horas_semanales,
           e.foto_perfil, o.nombre as obra_nombre, o.codigo as obra_codigo, o.direccion as obra_dir
    FROM registros r
    JOIN empleados e ON r.empleado_id=e.id
    LEFT JOIN obras o ON r.obra_id=o.id
    WHERE strftime('%Y-%m',r.fecha)=? AND r.anulado=0
    ${obraId ? 'AND r.obra_id='+obraId : ''}
    ORDER BY e.nombre, r.fecha
  `).all(mes);

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Construcciones FLP';
  const ws = wb.addWorksheet('Libro Asistencia '+mes, { pageSetup:{paperSize:9,orientation:'landscape'} });

  // Título
  ws.mergeCells('A1:N1');
  Object.assign(ws.getCell('A1'), {
    value: `CONSTRUCCIONES FLP — LIBRO DE ASISTENCIA — ${m}/${anio}`,
    font:{bold:true,size:13,color:{argb:'FFFFFFFF'}},
    fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FF1D3A6B'}},
    alignment:{horizontal:'center',vertical:'middle'}
  });
  ws.getRow(1).height = 28;

  // Subtítulo legal
  ws.mergeCells('A2:N2');
  Object.assign(ws.getCell('A2'), {
    value: 'Registro según Art. 33 Código del Trabajo — Dirección del Trabajo Chile',
    font:{italic:true,size:10,color:{argb:'FF1D3A6B'}},
    alignment:{horizontal:'center'}
  });

  // Encabezados
  const hdrs = ['RUT','Nombre','Cargo','Obra','Fecha','Día','H.Esperada','Entrada','Salida','H.Ordinarias','H.Extra','Atraso(min)','Estado','Foto'];
  ws.addRow(hdrs);
  ws.getRow(3).eachCell(c => {
    c.font={bold:true,color:{argb:'FFFFFFFF'}};
    c.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF1D3A6B'}};
    c.alignment={horizontal:'center',vertical:'middle'};
  });
  ws.getRow(3).height = 18;
  ws.columns=[
    {width:14},{width:22},{width:16},{width:16},{width:11},{width:8},
    {width:11},{width:10},{width:10},{width:12},{width:10},{width:13},{width:15},{width:10}
  ];

  const dias = ['Dom','Lun','Mar','Mié','Jue','Vie','Sáb'];
  const eTxt = {presente:'Presente',tardanza_leve:'Tardanza leve',tardanza_grave:'Tardanza grave',ausente:'Ausente'};
  const eCol = {presente:'FF2DC653',tardanza_leve:'FFFFA500',tardanza_grave:'FFFF4444',ausente:'FFAAAAAA'};

  let row = 4;
  let totalOrd = 0, totalExt = 0;

  for (const r of registros) {
    const wr = ws.getRow(row);
    const d = new Date(r.fecha+'T12:00:00');
    wr.getCell(1).value = formatRut(r.rut_numero, r.rut_dv);
    wr.getCell(2).value = r.nombre;
    wr.getCell(3).value = r.cargo;
    wr.getCell(4).value = r.obra_nombre||'—';
    wr.getCell(5).value = r.fecha;
    wr.getCell(6).value = dias[d.getDay()];
    wr.getCell(7).value = r.hora_esp;
    wr.getCell(8).value = r.hora_entrada||'—';
    wr.getCell(9).value = r.hora_salida||'—';
    wr.getCell(10).value = r.horas_ordinarias||0;
    wr.getCell(11).value = r.horas_extra||0;
    wr.getCell(12).value = r.minutos_atraso||0;
    wr.getCell(13).value = eTxt[r.estado]||r.estado;
    wr.getCell(13).font = {color:{argb:eCol[r.estado]||'FF000000'},bold:true};
    wr.height = 50;
    totalOrd += r.horas_ordinarias||0;
    totalExt += r.horas_extra||0;

    const fotoFile = r.foto_entrada || r.foto_perfil;
    if (fotoFile) {
      const fp = path.join(UPLOADS_DIR, fotoFile);
      if (fs.existsSync(fp)) {
        try {
          const ext = path.extname(fotoFile).replace('.','').toLowerCase();
          const imgId = wb.addImage({filename:fp, extension:ext==='jpg'?'jpeg':ext});
          ws.addImage(imgId, {tl:{col:13,row:row-1},br:{col:14,row:row},editAs:'oneCell'});
        } catch(e) {}
      }
    }
    row++;
  }

  // Totales
  row++;
  ws.mergeCells(`A${row}:I${row}`);
  ws.getCell(`A${row}`).value = `TOTALES DEL MES: ${registros.length} registros`;
  ws.getCell(`A${row}`).font = {bold:true};
  ws.getCell(`J${row}`).value = Math.round(totalOrd*100)/100;
  ws.getCell(`J${row}`).font = {bold:true};
  ws.getCell(`K${row}`).value = Math.round(totalExt*100)/100;
  ws.getCell(`K${row}`).font = {bold:true,color:{argb:'FFFF4444'}};

  // Firma
  row += 2;
  ws.mergeCells(`A${row}:G${row}`);
  ws.getCell(`A${row}`).value = 'Firma Empleador: ________________________________';
  ws.mergeCells(`H${row}:N${row}`);
  ws.getCell(`H${row}`).value = 'Timbre Empresa: ________________________________';

  return await wb.xlsx.writeBuffer();
}

// ─── Email ────────────────────────────────────────────────────
async function enviarReporte(db, fecha, obra) {
  if (!process.env.GMAIL_USER||!process.env.GMAIL_APP_PASSWORD) { console.log('⚠ Gmail no configurado'); return; }
  const fechaStr = fecha||new Date().toISOString().slice(0,10);
  const mes = fechaStr.slice(0,7);
  const destinatario = obra?.encargado_email || process.env.ADMIN_EMAIL || process.env.GMAIL_USER;
  const obraNombre = obra?.nombre||'Todas las obras';
  const buf = await generarLibroMensual(db, mes, obra?.id||null);
  const regs = db.prepare(`SELECT estado,COUNT(*) as cnt FROM registros WHERE fecha=? ${obra?'AND obra_id='+obra.id:''} AND anulado=0 GROUP BY estado`).all(fechaStr);
  const total = db.prepare(`SELECT COUNT(*) as cnt FROM empleados WHERE activo=1 ${obra?'AND obra_id='+obra.id:''}`).get().cnt;
  await nodemailer.createTransport({service:'gmail',auth:{user:process.env.GMAIL_USER,pass:process.env.GMAIL_APP_PASSWORD}})
    .sendMail({
      from:`"Asistencia FLP" <${process.env.GMAIL_USER}>`,
      to: destinatario,
      subject:`📋 Asistencia ${obraNombre} — ${fechaStr}`,
      html:`<div style="font-family:Arial;max-width:600px;">
        <div style="background:#1D3A6B;color:white;padding:20px;border-bottom:3px solid #C9A84C;border-radius:8px 8px 0 0;">
          <h2 style="margin:0;">Reporte Diario — ${obraNombre}</h2>
          <p style="margin:4px 0 0;opacity:.8;">${fechaStr}</p></div>
        <div style="padding:20px;background:#f8f9fa;border:1px solid #dee2e6;border-radius:0 0 8px 8px;">
          <p>Registrados: <strong>${regs.reduce((a,b)=>a+b.cnt,0)} / ${total}</strong></p>
          <p>${regs.map(r=>`${r.estado}: ${r.cnt}`).join(' | ')||'Sin registros'}</p>
          <p style="color:#666;font-size:13px;">Libro mensual adjunto con todos los registros del mes.</p>
        </div></div>`,
      attachments:[{filename:`libro_asistencia_${obraNombre.replace(/\s+/g,'_')}_${mes}.xlsx`,content:buf,
        contentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}]
    });
  console.log(`✅ Reporte → ${destinatario}`);
}

// ─── Iniciar ──────────────────────────────────────────────────
initDb().then(db => {

  // Cron 19:00 lun-sáb: reporte + respaldo
  cron.schedule('0 19 * * 1-6', async () => {
    hacerRespaldo(db);
    const obras = db.prepare('SELECT * FROM obras WHERE activa=1').all();
    for (const o of obras) await enviarReporte(db, null, o);
  }, { timezone:'America/Santiago' });

  // Respaldo diario a las 23:00
  cron.schedule('0 23 * * *', () => hacerRespaldo(db), { timezone:'America/Santiago' });

  // ── OBRAS ─────────────────────────────────────────────────
  app.get('/api/obras', authMw, (req,res) => {
    res.json(db.prepare('SELECT o.*,COUNT(e.id) as total_empleados FROM obras o LEFT JOIN empleados e ON e.obra_id=o.id AND e.activo=1 GROUP BY o.id ORDER BY o.nombre').all());
  });
  app.post('/api/obras', authMw, soloAdmin, (req,res) => {
    const {nombre,codigo,cliente,direccion,encargado_email}=req.body;
    if (!nombre||!codigo) return res.status(400).json({error:'Nombre y código requeridos'});
    try {
      const r=db.prepare('INSERT INTO obras (nombre,codigo,cliente,direccion,encargado_email) VALUES (?,?,?,?,?)').run(nombre,codigo,cliente||null,direccion||null,encargado_email||null);
      log(db,req.user.email,'admin','crear_obra',nombre,req.ip);
      res.json({id:r.lastInsertRowid,nombre,codigo});
    } catch(e){res.status(400).json({error:'El código ya existe'});}
  });
  app.put('/api/obras/:id', authMw, soloAdmin, (req,res) => {
    const {nombre,codigo,cliente,direccion,encargado_email,activa}=req.body;
    const o=db.prepare('SELECT * FROM obras WHERE id=?').get(req.params.id);
    if (!o) return res.status(404).json({error:'No encontrada'});
    db.prepare('UPDATE obras SET nombre=?,codigo=?,cliente=?,direccion=?,encargado_email=?,activa=? WHERE id=?')
      .run(nombre||o.nombre,codigo||o.codigo,cliente||o.cliente,direccion||o.direccion,encargado_email||o.encargado_email,activa!==undefined?activa:o.activa,req.params.id);
    log(db,req.user.email,'admin','editar_obra',nombre,req.ip);
    res.json({ok:true});
  });
  app.delete('/api/obras/:id', authMw, soloAdmin, (req,res) => {
    db.prepare('UPDATE obras SET activa=0 WHERE id=?').run(req.params.id);
    res.json({ok:true});
  });

  // ── AUTH ──────────────────────────────────────────────────
  app.post('/api/auth/admin', (req,res) => {
    const {email,password}=req.body;
    const a=db.prepare('SELECT * FROM admin WHERE email=?').get(email);
    if (!a||!bcrypt.compareSync(password,a.password_hash)) {
      log(db,email,'admin','login_fallido','',req.ip);
      return res.status(401).json({error:'Credenciales incorrectas'});
    }
    const token=jwt.sign({id:a.id,nombre:a.nombre,email:a.email,rol:'admin',es_superadmin:a.es_superadmin===1},JWT_SECRET,{expiresIn:'12h'});
    log(db,email,'admin','login_ok','',req.ip);
    res.json({token,nombre:a.nombre,rol:'admin',es_superadmin:a.es_superadmin===1});
  });
  app.post('/api/auth/empleado', (req,res) => {
    const {email,password}=req.body;
    const e=db.prepare('SELECT * FROM empleados WHERE email=? AND activo=1').get(email);
    if (!e||!bcrypt.compareSync(password,e.password_hash)) {
      log(db,email,'empleado','login_fallido','',req.ip);
      return res.status(401).json({error:'Credenciales incorrectas'});
    }
    const token=jwt.sign({id:e.id,nombre:e.nombre,email:e.email,rol:'empleado',obra_id:e.obra_id},JWT_SECRET,{expiresIn:'12h'});
    log(db,email,'empleado','login_ok','',req.ip);
    res.json({token,nombre:e.nombre,rol:'empleado',id:e.id,obra_id:e.obra_id});
  });

  // ── EMPLEADOS ─────────────────────────────────────────────
  app.get('/api/empleados', authMw, soloAdmin, (req,res) => {
    const {obra_id}=req.query;
    let q='SELECT e.*,o.nombre as obra_nombre,o.codigo as obra_codigo FROM empleados e LEFT JOIN obras o ON e.obra_id=o.id WHERE e.activo=1';
    const p=[];
    if (obra_id){q+=' AND e.obra_id=?';p.push(obra_id);}
    res.json(db.prepare(q+' ORDER BY o.nombre,e.nombre').all(...p));
  });
  app.get('/api/empleados/me', authMw, (req,res) => {
    res.json(db.prepare('SELECT e.*,o.nombre as obra_nombre,o.codigo as obra_codigo FROM empleados e LEFT JOIN obras o ON e.obra_id=o.id WHERE e.id=?').get(req.user.id));
  });
  app.post('/api/empleados', authMw, soloAdmin, upload.single('foto'), (req,res) => {
    const {nombre,rut_numero,rut_dv,cargo,email,telefono,password,hora_entrada,hora_salida,horas_semanales,obra_id}=req.body;
    if (!nombre||!rut_numero||!rut_dv||!cargo||!password) return res.status(400).json({error:'Faltan campos obligatorios'});
    if (!validarRut(rut_numero,rut_dv)) return res.status(400).json({error:'RUT inválido. Verifique el número y dígito verificador'});
    const foto=req.file?req.file.filename:null;
    try {
      const r=db.prepare('INSERT INTO empleados (nombre,rut_numero,rut_dv,cargo,email,telefono,password_hash,hora_entrada,hora_salida,horas_semanales,foto_perfil,obra_id) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)')
        .run(nombre,rut_numero.replace(/\D/g,''),rut_dv.toUpperCase(),cargo,email||null,telefono||null,bcrypt.hashSync(password,10),hora_entrada||'08:00',hora_salida||'17:00',horas_semanales||45,foto,obra_id||null);
      log(db,req.user.email,'admin','crear_empleado',nombre,req.ip);
      res.json({id:r.lastInsertRowid,nombre});
    } catch(e){res.status(400).json({error:'El email ya existe'});}
  });
  app.put('/api/empleados/:id', authMw, soloAdmin, upload.single('foto'), (req,res) => {
    const {nombre,rut_numero,rut_dv,cargo,email,telefono,hora_entrada,hora_salida,horas_semanales,password,activo,obra_id}=req.body;
    const e=db.prepare('SELECT * FROM empleados WHERE id=?').get(req.params.id);
    if (!e) return res.status(404).json({error:'No encontrado'});
    if (rut_numero&&rut_dv&&!validarRut(rut_numero,rut_dv)) return res.status(400).json({error:'RUT inválido'});
    const hash=password?bcrypt.hashSync(password,10):e.password_hash;
    const foto=req.file?req.file.filename:e.foto_perfil;
    db.prepare('UPDATE empleados SET nombre=?,rut_numero=?,rut_dv=?,cargo=?,email=?,telefono=?,hora_entrada=?,hora_salida=?,horas_semanales=?,password_hash=?,foto_perfil=?,activo=?,obra_id=? WHERE id=?')
      .run(nombre||e.nombre,rut_numero?rut_numero.replace(/\D/g,''):e.rut_numero,rut_dv?rut_dv.toUpperCase():e.rut_dv,cargo||e.cargo,email||e.email,telefono||e.telefono,hora_entrada||e.hora_entrada,hora_salida||e.hora_salida,horas_semanales||e.horas_semanales,hash,foto,activo!==undefined?activo:e.activo,obra_id||e.obra_id,req.params.id);
    log(db,req.user.email,'admin','editar_empleado',nombre||e.nombre,req.ip);
    res.json({ok:true});
  });
  app.delete('/api/empleados/:id', authMw, soloAdmin, (req,res) => {
    db.prepare('UPDATE empleados SET activo=0 WHERE id=?').run(req.params.id);
    log(db,req.user.email,'admin','desactivar_empleado',req.params.id,req.ip);
    res.json({ok:true});
  });

  // ── REGISTROS ─────────────────────────────────────────────
  app.post('/api/registros/entrada', authMw, (req,res) => {
    const {foto_base64,latitud,longitud,empleado_id}=req.body;
    const empId=req.user.rol==='admin'?empleado_id:req.user.id;
    const fecha=new Date().toISOString().slice(0,10);
    const hora=new Date().toLocaleTimeString('es-CL',{hour:'2-digit',minute:'2-digit'});
    const emp=db.prepare('SELECT * FROM empleados WHERE id=?').get(empId);
    if (!emp) return res.status(404).json({error:'Empleado no encontrado'});
    if (db.prepare('SELECT id FROM registros WHERE empleado_id=? AND fecha=? AND anulado=0').get(empId,fecha))
      return res.status(400).json({error:'Ya tienes entrada registrada hoy'});
    const foto=guardarFoto(foto_base64);
    const {estado,minutos}=calcEstado(hora,emp.hora_entrada);
    db.prepare('INSERT INTO registros (empleado_id,obra_id,fecha,hora_entrada,foto_entrada,lat_entrada,lon_entrada,estado,minutos_atraso) VALUES (?,?,?,?,?,?,?,?,?)')
      .run(empId,emp.obra_id,fecha,hora,foto,latitud||null,longitud||null,estado,minutos);
    log(db,emp.nombre,'empleado','entrada',`${hora} ${estado}`,req.ip);
    res.json({ok:true,hora,estado,minutos_atraso:minutos});
  });

  app.post('/api/registros/salida', authMw, (req,res) => {
    const {foto_base64,latitud,longitud,empleado_id}=req.body;
    const empId=req.user.rol==='admin'?empleado_id:req.user.id;
    const fecha=new Date().toISOString().slice(0,10);
    const hora=new Date().toLocaleTimeString('es-CL',{hour:'2-digit',minute:'2-digit'});
    const reg=db.prepare('SELECT * FROM registros WHERE empleado_id=? AND fecha=? AND anulado=0').get(empId,fecha);
    if (!reg) return res.status(400).json({error:'Sin entrada registrada hoy'});
    if (reg.hora_salida) return res.status(400).json({error:'Salida ya registrada'});
    const emp=db.prepare('SELECT * FROM empleados WHERE id=?').get(empId);
    const foto=guardarFoto(foto_base64);
    const {ordinarias,extra}=calcHorasMinutos(reg.hora_entrada,hora,emp?.horas_semanales);
    db.prepare('UPDATE registros SET hora_salida=?,foto_salida=?,lat_salida=?,lon_salida=?,horas_ordinarias=?,horas_extra=? WHERE id=?')
      .run(hora,foto,latitud||null,longitud||null,ordinarias,extra,reg.id);
    log(db,emp?.nombre||empId,'empleado','salida',hora,req.ip);
    res.json({ok:true,hora,horas_ordinarias:ordinarias,horas_extra:extra});
  });

  // Corrección de registro (solo superadmin, con motivo obligatorio)
  app.put('/api/registros/:id/corregir', authMw, soloSuperAdmin, (req,res) => {
    const {hora_entrada,hora_salida,motivo}=req.body;
    if (!motivo||motivo.trim().length<10) return res.status(400).json({error:'El motivo es obligatorio y debe tener al menos 10 caracteres'});
    const reg=db.prepare('SELECT * FROM registros WHERE id=?').get(req.params.id);
    if (!reg) return res.status(404).json({error:'Registro no encontrado'});
    if (reg.anulado) return res.status(400).json({error:'No se puede corregir un registro anulado'});
    const emp=db.prepare('SELECT * FROM empleados WHERE id=?').get(reg.empleado_id);

    // Guardar correcciones en historial
    if (hora_entrada && hora_entrada !== reg.hora_entrada) {
      db.prepare('INSERT INTO correcciones (registro_id,empleado_id,campo_modificado,valor_anterior,valor_nuevo,motivo,modificado_por) VALUES (?,?,?,?,?,?,?)')
        .run(reg.id,reg.empleado_id,'hora_entrada',reg.hora_entrada,hora_entrada,motivo,req.user.email);
    }
    if (hora_salida && hora_salida !== reg.hora_salida) {
      db.prepare('INSERT INTO correcciones (registro_id,empleado_id,campo_modificado,valor_anterior,valor_nuevo,motivo,modificado_por) VALUES (?,?,?,?,?,?,?)')
        .run(reg.id,reg.empleado_id,'hora_salida',reg.hora_salida,hora_salida,motivo,req.user.email);
    }

    const nuevaEntrada = hora_entrada||reg.hora_entrada;
    const nuevaSalida = hora_salida||reg.hora_salida;
    const {estado,minutos}=calcEstado(nuevaEntrada,emp?.hora_entrada||'08:00');
    const {ordinarias,extra}=calcHorasMinutos(nuevaEntrada,nuevaSalida,emp?.horas_semanales);

    db.prepare('UPDATE registros SET hora_entrada=?,hora_salida=?,estado=?,minutos_atraso=?,horas_ordinarias=?,horas_extra=? WHERE id=?')
      .run(nuevaEntrada,nuevaSalida,estado,minutos,ordinarias,extra,reg.id);

    log(db,req.user.email,'admin','correccion_registro',`ID:${reg.id} motivo:${motivo}`,req.ip);
    res.json({ok:true,mensaje:'Corrección registrada con historial'});
  });

  // Anular registro (nunca borrar, solo marcar)
  app.put('/api/registros/:id/anular', authMw, soloSuperAdmin, (req,res) => {
    const {motivo}=req.body;
    if (!motivo||motivo.trim().length<10) return res.status(400).json({error:'El motivo es obligatorio (mín. 10 caracteres)'});
    const reg=db.prepare('SELECT * FROM registros WHERE id=?').get(req.params.id);
    if (!reg) return res.status(404).json({error:'Registro no encontrado'});
    if (reg.anulado) return res.status(400).json({error:'Ya está anulado'});
    db.prepare('UPDATE registros SET anulado=1,motivo_anulacion=?,anulado_por=?,anulado_en=datetime("now","localtime") WHERE id=?')
      .run(motivo,req.user.email,reg.id);
    log(db,req.user.email,'admin','anular_registro',`ID:${reg.id} motivo:${motivo}`,req.ip);
    res.json({ok:true});
  });

  // Historial de correcciones de un registro
  app.get('/api/registros/:id/historial', authMw, soloAdmin, (req,res) => {
    res.json(db.prepare('SELECT * FROM correcciones WHERE registro_id=? ORDER BY modificado_en DESC').all(req.params.id));
  });

  app.get('/api/registros', authMw, (req,res) => {
    const {fecha,mes,obra_id,incluir_anulados}=req.query;
    const anulFiltro = incluir_anulados==='1' ? '' : 'AND r.anulado=0';
    let q,p;
    if (req.user.rol==='admin') {
      const base=`SELECT r.*,e.nombre,e.rut_numero,e.rut_dv,e.cargo,e.foto_perfil,o.nombre as obra_nombre,o.codigo as obra_codigo FROM registros r JOIN empleados e ON r.empleado_id=e.id LEFT JOIN obras o ON r.obra_id=o.id`;
      if (fecha) {
        q=base+` WHERE r.fecha=? ${anulFiltro} ${obra_id?'AND r.obra_id='+obra_id:''} ORDER BY o.nombre,e.nombre`;
        p=[fecha];
      } else if (mes) {
        q=base+` WHERE strftime('%Y-%m',r.fecha)=? ${anulFiltro} ${obra_id?'AND r.obra_id='+obra_id:''} ORDER BY r.fecha DESC,e.nombre`;
        p=[mes];
      } else {
        const h=new Date().toISOString().slice(0,10);
        q=base+` WHERE r.fecha=? ${anulFiltro} ${obra_id?'AND r.obra_id='+obra_id:''} ORDER BY o.nombre,e.nombre`;
        p=[h];
      }
    } else {
      if (mes){q=`SELECT * FROM registros WHERE empleado_id=? AND strftime('%Y-%m',fecha)=? ${anulFiltro} ORDER BY fecha DESC`;p=[req.user.id,mes];}
      else{q=`SELECT * FROM registros WHERE empleado_id=? ${anulFiltro} ORDER BY fecha DESC LIMIT 90`;p=[req.user.id];}
    }
    res.json(db.prepare(q).all(...p));
  });

  // ── STATS ─────────────────────────────────────────────────
  app.get('/api/stats/hoy', authMw, soloAdmin, (req,res) => {
    const {obra_id}=req.query;
    const hoy=new Date().toISOString().slice(0,10);
    const wo=obra_id?' AND obra_id='+obra_id:'';
    const we=obra_id?' AND obra_id='+obra_id:'';
    const total=db.prepare(`SELECT COUNT(*) as cnt FROM empleados WHERE activo=1${we}`).get().cnt;
    const regs=db.prepare(`SELECT estado,COUNT(*) as cnt FROM registros WHERE fecha=? AND anulado=0${wo} GROUP BY estado`).all(hoy);
    const conReg=db.prepare(`SELECT COUNT(*) as cnt FROM registros WHERE fecha=? AND anulado=0${wo}`).get(hoy).cnt;
    const porObra=db.prepare(`SELECT o.nombre,o.codigo,COUNT(r.id) as presentes,COUNT(e.id) as total_emp
      FROM obras o LEFT JOIN empleados e ON e.obra_id=o.id AND e.activo=1
      LEFT JOIN registros r ON r.empleado_id=e.id AND r.fecha=? AND r.anulado=0 AND (r.estado='presente' OR r.estado LIKE 'tardanza%')
      WHERE o.activa=1 GROUP BY o.id ORDER BY o.nombre`).all(hoy);
    res.json({total,ausentes:total-conReg,registros:regs,conRegistro:conReg,porObra});
  });

  // ── REPORTES ──────────────────────────────────────────────
  app.get('/api/reportes/libro', authMw, soloAdmin, async (req,res) => {
    const mes=req.query.mes||new Date().toISOString().slice(0,7);
    const obraId=req.query.obra_id||null;
    try {
      const buf=await generarLibroMensual(db,mes,obraId);
      log(db,req.user.email,'admin','exportar_libro',`mes:${mes}`,req.ip);
      res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition',`attachment; filename=libro_asistencia_${mes}.xlsx`);
      res.send(buf);
    } catch(e){res.status(500).json({error:e.message});}
  });

  app.post('/api/reportes/enviar', authMw, soloAdmin, async (req,res) => {
    const {fecha,obra_id}=req.body;
    try {
      if (obra_id) { const o=db.prepare('SELECT * FROM obras WHERE id=?').get(obra_id); await enviarReporte(db,fecha,o); }
      else { const obras=db.prepare('SELECT * FROM obras WHERE activa=1').all(); for (const o of obras) await enviarReporte(db,fecha,o); }
      res.json({ok:true});
    } catch(e){res.status(500).json({error:e.message});}
  });

  // Respaldo manual
  app.post('/api/respaldo', authMw, soloSuperAdmin, (req,res) => {
    hacerRespaldo(db);
    res.json({ok:true,mensaje:'Respaldo creado'});
  });
  app.get('/api/respaldos', authMw, soloSuperAdmin, (req,res) => {
    res.json(db.prepare('SELECT * FROM respaldos ORDER BY creado_en DESC LIMIT 30').all());
  });

  // ── LOGS / FISCALIZADOR ───────────────────────────────────
  app.get('/api/logs', authMw, soloAdmin, (req,res) => {
    res.json(db.prepare('SELECT * FROM logs ORDER BY fecha DESC LIMIT 200').all());
  });

  // Vista pública fiscalizador (token especial)
  app.get('/api/fiscalizador', (req,res) => {
    const key=req.query.key;
    if (!key||key!==process.env.FISCALIZADOR_KEY) return res.status(403).json({error:'Acceso no autorizado'});
    const mes=req.query.mes||new Date().toISOString().slice(0,7);
    const obra_id=req.query.obra_id||null;
    const regs=db.prepare(`
      SELECT r.fecha,r.hora_entrada,r.hora_salida,r.estado,r.minutos_atraso,r.horas_ordinarias,r.horas_extra,r.anulado,
             e.nombre,e.rut_numero,e.rut_dv,e.cargo,o.nombre as obra,o.codigo as obra_codigo
      FROM registros r JOIN empleados e ON r.empleado_id=e.id LEFT JOIN obras o ON r.obra_id=o.id
      WHERE strftime('%Y-%m',r.fecha)=? ${obra_id?'AND r.obra_id='+obra_id:''}
      ORDER BY e.nombre,r.fecha
    `).all(mes);
    res.json({mes,generado:new Date().toISOString(),empresa:'Construcciones FLP',registros:regs.map(r=>({...r,rut:formatRut(r.rut_numero,r.rut_dv)}))});
  });

  app.get('*',(req,res)=>res.sendFile(path.join(__dirname,'public','index.html')));

  app.listen(PORT,()=>{
    console.log(`🚀 Servidor en http://localhost:${PORT}`);
    console.log(`📧 Gmail: ${process.env.GMAIL_USER||'⚠ No configurado'}`);
    console.log(`🔒 Fiscalizador key: ${process.env.FISCALIZADOR_KEY||'⚠ No configurada'}`);
  });
}).catch(err=>{console.error('Error DB:',err);process.exit(1);});
