const initSqlJs = require('sql.js');
const fs = require('fs');
const path = require('path');
const bcrypt = require('bcryptjs');

const DB_PATH = process.env.DB_PATH || path.join(__dirname, 'asistencia.db');
let _sqliteDb = null;
let _dbWrapper = null;

function saveToDisk() {
  if (_sqliteDb) {
    const data = _sqliteDb.export();
    fs.writeFileSync(DB_PATH, Buffer.from(data));
  }
}

function makeWrapper(sqliteDb) {
  return {
    exec(sql) { sqliteDb.run(sql); saveToDisk(); },
    pragma(p) { try { sqliteDb.run('PRAGMA ' + p); } catch(e) {} },
    prepare(sql) {
      return {
        run(...params) {
          sqliteDb.run(sql, params); saveToDisk();
          const r = sqliteDb.exec('SELECT last_insert_rowid() as id');
          return { lastInsertRowid: r[0]?.values[0]?.[0] || 0 };
        },
        get(...params) {
          const stmt = sqliteDb.prepare(sql); stmt.bind(params);
          if (stmt.step()) {
            const cols = stmt.getColumnNames(), vals = stmt.get(); stmt.free();
            const obj = {}; cols.forEach((c,i) => obj[c] = vals[i]); return obj;
          }
          stmt.free(); return undefined;
        },
        all(...params) {
          const results = [], stmt = sqliteDb.prepare(sql); stmt.bind(params);
          while (stmt.step()) {
            const cols = stmt.getColumnNames(), vals = stmt.get();
            const obj = {}; cols.forEach((c,i) => obj[c] = vals[i]); results.push(obj);
          }
          stmt.free(); return results;
        }
      };
    }
  };
}

async function initDb() {
  if (_dbWrapper) return _dbWrapper;
  const SQL = await initSqlJs();
  _sqliteDb = fs.existsSync(DB_PATH)
    ? new SQL.Database(fs.readFileSync(DB_PATH))
    : new SQL.Database();
  const db = makeWrapper(_sqliteDb);
  _dbWrapper = db;

  // OBRAS
  _sqliteDb.run(`CREATE TABLE IF NOT EXISTS obras (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL,
    codigo TEXT UNIQUE NOT NULL,
    cliente TEXT,
    direccion TEXT,
    encargado_email TEXT,
    activa INTEGER DEFAULT 1,
    creado_en TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // EMPLEADOS - con RUT separado
  _sqliteDb.run(`CREATE TABLE IF NOT EXISTS empleados (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL,
    rut_numero TEXT NOT NULL,
    rut_dv TEXT NOT NULL,
    cargo TEXT NOT NULL,
    email TEXT UNIQUE,
    telefono TEXT,
    password_hash TEXT NOT NULL,
    hora_entrada TEXT DEFAULT '08:00',
    hora_salida TEXT DEFAULT '17:00',
    horas_semanales INTEGER DEFAULT 45,
    foto_perfil TEXT,
    obra_id INTEGER,
    activo INTEGER DEFAULT 1,
    creado_en TEXT DEFAULT (datetime('now','localtime')),
    FOREIGN KEY (obra_id) REFERENCES obras(id)
  )`);

  // REGISTROS - inmutables, solo se anulan
  _sqliteDb.run(`CREATE TABLE IF NOT EXISTS registros (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    empleado_id INTEGER NOT NULL,
    obra_id INTEGER,
    fecha TEXT NOT NULL,
    hora_entrada TEXT,
    hora_salida TEXT,
    foto_entrada TEXT,
    foto_salida TEXT,
    lat_entrada REAL,
    lon_entrada REAL,
    lat_salida REAL,
    lon_salida REAL,
    estado TEXT DEFAULT 'presente',
    minutos_atraso INTEGER DEFAULT 0,
    horas_ordinarias REAL DEFAULT 0,
    horas_extra REAL DEFAULT 0,
    anulado INTEGER DEFAULT 0,
    motivo_anulacion TEXT,
    anulado_por TEXT,
    anulado_en TEXT,
    creado_en TEXT DEFAULT (datetime('now','localtime')),
    FOREIGN KEY (empleado_id) REFERENCES empleados(id),
    FOREIGN KEY (obra_id) REFERENCES obras(id)
  )`);

  // CORRECCIONES - historial de cambios con motivo obligatorio
  _sqliteDb.run(`CREATE TABLE IF NOT EXISTS correcciones (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    registro_id INTEGER NOT NULL,
    empleado_id INTEGER NOT NULL,
    campo_modificado TEXT NOT NULL,
    valor_anterior TEXT,
    valor_nuevo TEXT,
    motivo TEXT NOT NULL,
    modificado_por TEXT NOT NULL,
    modificado_en TEXT DEFAULT (datetime('now','localtime')),
    FOREIGN KEY (registro_id) REFERENCES registros(id)
  )`);

  // ADMIN
  _sqliteDb.run(`CREATE TABLE IF NOT EXISTS admin (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombre TEXT NOT NULL,
    email TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    es_superadmin INTEGER DEFAULT 0,
    creado_en TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // LOGS DE ACCESO Y ACCIONES
  _sqliteDb.run(`CREATE TABLE IF NOT EXISTS logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    usuario TEXT,
    rol TEXT,
    accion TEXT,
    detalle TEXT,
    ip TEXT,
    fecha TEXT DEFAULT (datetime('now','localtime'))
  )`);

  // RESPALDOS
  _sqliteDb.run(`CREATE TABLE IF NOT EXISTS respaldos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    archivo TEXT,
    tamanio INTEGER,
    creado_en TEXT DEFAULT (datetime('now','localtime'))
  )`);

  saveToDisk();

  // Superadmin por defecto
  if (!db.prepare('SELECT id FROM admin WHERE email=?').get('admin@flp.cl')) {
    db.prepare('INSERT INTO admin (nombre,email,password_hash,es_superadmin) VALUES (?,?,?,1)')
      .run('Administrador FLP','admin@flp.cl',bcrypt.hashSync('admin123',10));
    console.log('✅ Superadmin: admin@flp.cl / admin123 (¡cámbiala!)');
  }

  // Obras demo
  if (!db.prepare('SELECT id FROM obras LIMIT 1').get()) {
    const so = db.prepare('INSERT INTO obras (nombre,codigo,cliente,direccion,encargado_email) VALUES (?,?,?,?,?)');
    so.run('Edificio Central','OBR-001','Cliente A','Av. Principal 123, Santiago','jefe1@flp.cl');
    so.run('Planta Norte','OBR-002','Cliente B','Ruta 5 Norte Km 10','jefe2@flp.cl');
    console.log('✅ Obras demo creadas');
  }

  // Empleados demo
  if (!db.prepare('SELECT id FROM empleados LIMIT 1').get()) {
    const h = bcrypt.hashSync('empleado123',10);
    const se = db.prepare('INSERT INTO empleados (nombre,rut_numero,rut_dv,cargo,email,password_hash,hora_entrada,hora_salida,obra_id) VALUES (?,?,?,?,?,?,?,?,?)');
    se.run('Ana Torres','12345678','9','Coordinadora','ana@flp.cl',h,'08:00','17:00',1);
    se.run('Luis Ramos','9876543','K','Técnico','luis@flp.cl',h,'08:00','17:00',1);
    se.run('Camila Vega','11223344','5','Diseñadora','camila@flp.cl',h,'08:00','17:00',2);
    console.log('✅ Empleados demo: contraseña empleado123');
  }

  return db;
}

module.exports = { initDb, saveToDisk, getDb: () => _dbWrapper };
