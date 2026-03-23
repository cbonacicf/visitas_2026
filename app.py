#!/usr/bin/env python
# coding: utf-8

import polars as pl
from datetime import datetime, date, time, timedelta
import os
import json
import pytz
import io
from io import BytesIO
import base64
from collections import namedtuple

from sqlalchemy import create_engine, URL, text
import psycopg2
from sqlalchemy.orm import Session
from sqlalchemy.ext.automap import automap_base
from sqlalchemy.pool import NullPool

import dash
import dash_ag_grid as dag
import dash_bootstrap_components as dbc
from dash import html, dcc, Input, Output, State, Dash
from dash.exceptions import PreventUpdate

import matplotlib
matplotlib.use('agg')
import matplotlib.pyplot as plt

from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.colors import HexColor

# import numpy as np
import pickle

### Parámetros
fecha_inicial = date(2026, 3, 1)
fecha_final = date(2026, 11, 30)
fecha_limite = date(2026, 12, 31)

usuario = 0

### Conexión
# objeto_url = URL.create(
#     'postgresql+psycopg2',
#     username = os.environ['PGUSER'],
#     password = os.environ['PGPASSWORD'],
#     host = os.environ['PGHOST'],
#     port = os.environ['PGPORT'],
#     database = os.environ['PGDATABASE'],
# )

engine = create_engine(os.environ['DATABASE_URL'], pool_pre_ping=True) # poolclass=NullPool)

# engine = create_engine(objeto_url, pool_pre_ping=True) #, poolclass=NullPool)

### Lectura de insumos
universidades_excluidas = [27, 38]
with open('./data/universidades.json', 'r') as f:
    universidades = json.load(f)

universidades = {int(k): v for k, v in universidades.items() if int(k) not in universidades_excluidas}
usuarios = {0: 'Usuario no acreditado'} | universidades

colegios = dict(
    pl.read_parquet('./data/colegios.parquet')
    .select(['rbd', 'nombre'])
    .rows()
)

colegios_comuna = dict(
    pl.read_parquet('./data/colegios.parquet')
    .select(['rbd', 'cod_com'])
    .rows()
)

feriados = (
    pl.read_parquet('./data/feriados2026.parquet')
    .to_series()
    .to_list()
)

comunas = dict(
    pl.read_parquet('./data/comunas.parquet')
    .sort('comuna')
    .rows()
)

horas_15 = dict(
    pl.read_parquet('./data/div_horas.parquet')
    .rows()
)

### Traduce a español
dias_es = {
    'Monday': 'Lunes',
    'Tuesday': 'Martes',
    'Wednesday': 'Miércoles',
    'Thursday': 'Jueves',
    'Friday': 'Viernes',
    'Saturday': 'Sábado',
    'Sunday': 'Domingo',
}

meses_es = {
    'January': 'de Enero de',
    'February': 'de Febrero de',
    'March': 'de Marzo de',
    'April': 'de Abril de',
    'May': 'de Mayo de',
    'June': 'de Junio de',
    'July': 'de Julio de',
    'August': 'de Agosto de',
    'September': 'de Septiembre de',
    'October': 'de Octubre de',
    'November': 'de Noviembre de',
    'December': 'de Diciembre de',
}

items_es = dias_es | meses_es
tuplas_es = items_es.items()

### Funciones
inv = lambda dic: {v: k for k, v in dic.items()}
a_fecha = lambda fecha: datetime.strptime(fecha, '%Y-%m-%d').date()

def a_hora(hora):
    if bool(hora):
        return datetime.strptime(hora, '%H:%M:%S').time()
    else:
        return None

ahora = lambda: datetime.now(pytz.timezone('America/Santiago')).date() # + timedelta(days=38)
sig_laboral = lambda fecha=ahora(), dif=0: sorted(list({fecha + timedelta(days=i) for i in range(14)}.difference(feriados)))[dif]
dia_laboral = lambda: max(sig_laboral(fecha_inicial), sig_laboral())
fn_mes = lambda: dia_laboral().month

def opciones(dic):
    return [{'label': v, 'value': k} for k, v in dic.items()]

dic_meses = dict(zip(range(1, 13), ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']))

def opciones_meses():
    mi = fecha_inicial.month
    ma = fecha_final.month
    map_meses = {0: 'Todos'} | {k: v for k, v in dic_meses.items() if mi <= k <= ma}
    return opciones(map_meses)

def convierte_hora(hora):
    if hora is None:
        hora = '00:00:00'

    return datetime.strptime(hora, '%H:%M:%S').time()

def convierte_esp(fecha):
    for ingles, espanol in tuplas_es:
        fecha = fecha.replace(ingles, espanol)
    return fecha

binario = lambda x: int(bool(x))

dic_fn_modifica = {
    'fecha': a_fecha,
    'hora_ini': a_hora,
    'hora_ter': a_hora,
    'hora_ins': a_hora,
    'fecha_lim': a_fecha,
}

### Orden
orden_visita = ['programada_id', 'orden', 'fecha', 'rbd', 'nombre', 'organizador_id', 'organizador', 'estatus']
orden_fecha = ['orden', 'fecha', 'rbd', 'nombre', 'organizador', 'estatus']
orden_modifica = ['programada_id', 'orden', 'organizador_id', 'fecha', 'rbd', 'nombre', 'estatus', 'fecha_lim']
orden_invita = ['programada_id', 'orden', 'fecha', 'rbd', 'nombre', 'direccion', 'comuna_id', 'estatus']

orden_extendido = ['programada_id', 'orden', 'fecha', 'organizador_id', 'organizador', 'nombre', 'rbd', 'direccion', 'comuna_id', 'hora_ins', 'hora_ini', 'hora_fin',
                   'contacto', 'contacto_tel', 'contacto_mail', 'contacto_cargo', 'orientador', 'orientador_tel', 'orientador_mail', 'asistentes', 'asistentes_prog',
                   'estatus', 'observaciones']

map_orden_todas = {
    'programada_id': 'ID',
    'orden': 'orden',
    'fecha': 'Fecha',
    'rbd': 'RBD',
    'nombre': 'Colegio',
    'organizador_id': 'Código',
    'organizador': 'Organizador',
    'direccion': 'Dirección',
    'comuna_id': 'Comuna',
    'hora_ini': 'Inicio',
    'hora_fin': 'Término',
    'hora_ins': 'Instalación',
    'contacto': 'Contacto',
    'contacto_tel': 'Teléfono contacto',
    'contacto_mail': 'Correo contacto',
    'contacto_cargo': 'Cargo contacto',
    'orientador': 'Orientador',
    'orientador_tel': 'Teléfono orientador',
    'orientador_mail': 'Correo orientador',
    'asistentes': 'Asistentes estimados',
    'asistentes_prog': 'Programación de asistencia',
    'estatus': 'Estatus',
    'observaciones': 'Observaciones',
    'fecha_lim': 'fecha límite',
}

fn_orden = lambda orden: {k: map_orden_todas[k] for k in orden}

### Lectura de datos
Base = automap_base()
Base.prepare(autoload_with=engine)

Visita = Base.classes.programadas
Asiste = Base.classes.asisten

with open("./data/schema_programadas.pkl", "rb") as f:
    schema_programada = pickle.load(f)

with open("./data/schema_asisten.pkl", "rb") as f:
    schema_asisten = pickle.load(f)

actualiza_esquema = {
    'fecha': pl.Utf8,
    'hora_ini': pl.Utf8,
    'hora_fin': pl.Utf8,
    'hora_ins': pl.Utf8,
    'fecha_lim': pl.Utf8,
}

schema_programada_op = dict(schema_programada)
schema_programada_op.update(actualiza_esquema)
schema_programada_op = pl.Schema(schema_programada_op)

def lectura(db_nombre, db_schema):
    return pl.read_database(
        query = text(f'SELECT * FROM {db_nombre}'),
        connection = engine,
        schema_overrides=db_schema,
    )

def lectura_lazy(db_nombre, db_schema):
    return lectura(db_nombre, db_schema).lazy()

def obtiene_bloqueadas():
    sql = text("SELECT * FROM bloqueadas()")
    with Session(engine) as session:
        resultado = session.execute(sql).all()
    return [item[0] for item in resultado]

def fecha_str(fecha):
    if isinstance(fecha, str):
        fecha = datetime.strptime(fecha, '%Y-%m-%d').date()
    return fecha

def excluye_fecha(fecha, bloqueadas_loc):
    if fecha in bloqueadas_loc:
        bloqueadas_loc.remove(fecha)
    return bloqueadas_loc

def en_bloqueadas(fecha, excluye=None):
    fecha = fecha_str(fecha)
    bloqueadas_loc = obtiene_bloqueadas()
    if excluye:
        excluye = fecha_str(excluye)
        return (fecha in excluye_fecha(excluye, bloqueadas_loc))
    else:
        return (fecha in obtiene_bloqueadas())

def en_bloqueadas2(fecha, lista_bloqueadas, excluye=None):
    fecha = fecha_str(fecha)
    bloqueadas_loc = [fecha_str(fecha) for fecha in lista_bloqueadas]
    if excluye:
        excluye = fecha_str(excluye)
        return (fecha in excluye_fecha(excluye, bloqueadas_loc))
    else:
        return (fecha in bloqueadas_loc)

def fecha_a_str(df):
    return (df
        .with_columns(
            pl.col('fecha').dt.strftime('%Y-%m-%d'),
            pl.col('hora_ini').dt.strftime('%H:%M:%S'),
            pl.col('hora_fin').dt.strftime('%H:%M:%S'),
            pl.col('hora_ins').dt.strftime('%H:%M:%S'),
            pl.col('fecha_lim').dt.strftime('%Y-%m-%d'),
        )
    )

def str_a_fecha(df):
    return (df
        .with_columns(
            pl.col('fecha').str.strptime(pl.Date, '%Y-%m-%d'),
            pl.col('hora_ini').str.strptime(pl.Date, '%H:%M:%S'),
            pl.col('hora_fin').str.strptime(pl.Date, '%H:%M:%S'),
            pl.col('hora_ins').str.strptime(pl.Date, '%H:%M:%S'),
            pl.col('fecha_lim').str.strptime(pl.Date, '%Y-%m-%d'),
        )
    )

### Clases
class Metodos:

    # seleccionas las visitas programadas por el usuario
    def programadas_usuario(self):
        crt = (pl.col('organizador_id') == self.usuario) & (pl.col('fecha_lim') >= ahora())
        df = (
            self.programadas()
            .select(orden_modifica)
            .filter(crt)
            .sort(['fecha', 'orden'])
            .with_columns(
                pl.arange(1, pl.len()+1).alias('orden_local'),
                pl.when(pl.col('fecha') >= ahora()).then(pl.lit(0)).otherwise(pl.lit(1)).alias('tipo'),
            )
            .with_columns(
                pl.when((pl.col('tipo') == 0) & (pl.col('estatus') == 'Suspendida')).then(pl.lit(3)).otherwise(pl.col('tipo')).alias('tipo')
            )
            .with_columns(
                pl.when((pl.col('tipo') == 1) & (pl.col('fecha_lim') != fecha_limite)).then(pl.lit(2)).otherwise(pl.col('tipo')).alias('tipo'),
            )
            .select(['programada_id', 'orden_local', 'rbd', 'nombre', 'fecha', 'fecha_lim', 'tipo', 'estatus'])
            .rename({'orden_local': 'orden'})
            .collect()
        )
        return df.to_dicts(), len(df.filter(pl.col('tipo') == 1))


    # selecciona las invitaciones, distintas de 'Suspendidas', según la situación de asistencia del usuario
    def invitaciones(self, condicion):  # tendría que ingresar como parámetro la condición de asistencia de interés (1, 0, 2)
        return (
            self.asisten()
            .select(['programada_id', 'codigo', 'asiste', 'invitacion'])
            .filter((pl.col('codigo') == self.usuario) & (pl.col('asiste') == condicion))
            .join(
                (
                    self.programadas()
                    .select(orden_invita)
                    .filter((pl.col('fecha') >= ahora()) & (pl.col('estatus').is_in(['Confirmada', 'Por confirmar'])))  # aquí filtra las 'Suspendidas'
                ),
                on='programada_id',
                how='inner'
            )
            .sort(['fecha', 'orden'])
            .with_columns(
                pl.col('comuna_id').replace_strict(comunas, return_dtype=pl.Utf8).alias('comuna'),
                pl.arange(1, pl.len()+1).alias('orden_local'),
            )
            .select(['programada_id', 'orden_local', 'fecha', 'rbd', 'nombre', 'direccion', 'comuna', 'invitacion'])
            .rename({'orden_local': 'orden'})
            .collect()
            .to_dicts()
        )


    # cantidad de visitas vigentes no confirmadas
    def cantidad_visitas(self):
        visitas = set(
            self.asisten()
            .select(['programada_id', 'codigo', 'asiste'])
            .filter((pl.col('codigo') == self.usuario) & (pl.col('asiste') == 2))
            .join(
                (
                    self.programadas()
                    .select(['programada_id', 'fecha', 'estatus'])
                    .filter((pl.col('fecha') >= ahora()) & (pl.col('estatus').is_in(['Confirmada', 'Por confirmar'])))
                ),
                on='programada_id',
                how='inner'
            )
            .collect()
            .get_column('programada_id')
            .to_list()
        )
        return len(visitas)

    # texto de la cinta de advertencia
    def texto_advertencia(self):
        n_inv = self.cantidad_visitas()
        n_vis = self.programadas_usuario()[1]
        textos = [
            f' {n_inv} invitaciones por confirmar asistencia', # invitaciones
            f' {n_vis} visitas sin confirmación de realización', # realizadas
        ]
        txt = ''
        match (binario(n_inv), binario(n_vis)):
            case (0, 0):
                txt = ''
            case (1, 0):
                txt = 'Advertencia: Restan' + textos[0].capitalize() + '.'
            case (0, 1):
                txt = 'Advertencia: Permanecen' + textos[1].capitalize() + '.'
            case (1, 1):
                txt = 'Advertencia: Restan' + textos[0].capitalize() + ' y' + textos[1] + '.'
        return txt

class DatosInicio(Metodos):

    def __init__(self, base, base_asisten, usuario):
        self.base = base
        self.base_asisten = base_asisten
        self.usuario = usuario


    def programadas(self): 
        return pl.LazyFrame(self.base, schema=schema_programada_op).pipe(str_a_fecha)


    def asisten(self):
        return pl.LazyFrame(self.base_asisten, schema=schema_asisten)

class Actualiza(Metodos):

    def __init__(self, parametros):
        self.usuario = parametros['usuario']
        self.fecha = parametros['fecha_seleccionada']
        self.fecha_modifica = parametros['fecha_mod_seleccionada']
        self.mes = parametros['mes']
        self.bloqueadas = obtiene_bloqueadas()


    def programadas(self):
        return lectura_lazy('programadas', schema_programada)


    def asisten(self):
        return lectura_lazy('asisten', schema_asisten)


    def programadas_dic(self):
        return self.programadas().pipe(fecha_a_str).collect().to_dicts()


    def asisten_dic(self):
        return self.asisten().collect().to_dicts()


    # información de las visitas programadas por mes (visitas)
    def programadas_visita(self):
        df = self.programadas().select(orden_visita)
        if self.mes != 0:
            df = (
                df
                .filter(pl.col('fecha').dt.month() == self.mes)
                .with_columns(
                    pl.col('fecha').dt.strftime('%Y-%m-%d'),
                )
            )
    
        return df.sort(['fecha', 'orden']).collect().to_dicts()


    # datos de las visitas programadas para una determinada fecha (agrega): No tiene argumentos
    def programadas_fecha(self, fecha=None):
        if fecha == None:
            fecha = fecha_str(self.fecha)
        else:
            fecha = fecha_str(fecha)

        return (
            self.programadas()
            .select(orden_fecha)
            .filter((pl.col('fecha') == fecha) & (pl.col('estatus').is_in(['Confirmada', 'Por confirmar'])))
            .sort(['fecha', 'orden'])
            .with_columns(
                pl.arange(1, pl.len()+1).alias('orden_local'),
            )
            .select(['orden_local', 'rbd', 'nombre', 'organizador', 'estatus'])
            .rename({'orden_local': 'orden'})
            .collect()
            .to_dicts()
        )


    def oculta_advertencia(self):
        return (self.cantidad_visitas(), self.programadas_usuario()[1]) == (0, 0)  # cambio aquí


    def universidades_asisten(self, visita):
        return dict(
            self.asisten()
            .filter(pl.col('programada_id') == visita)
            .sort(['asiste', 'codigo'])
            .group_by('asiste', maintain_order=True)
            .agg(
                pl.col('codigo')
            )
            .collect()
            .rows()
        )

class DatosProgramadas:

    def __init__(self):
        self.ori, self.tra = self.lectura_trans()
        self.ori_dic, self.tra_dic = self.fn_dic()


    def lectura_trans(self):
        df = lectura_lazy('programadas', schema_programada)
        df_trans = fecha_a_str(df)
        return df, df_trans

    
    def fn_dic(self):
        return (item.collect().to_dicts() for item in self.lectura_trans())

class DatosAsisten:

    def __init__(self):
        self.ori = lectura_lazy('asisten', schema_asisten)
        self.ori_dic = self.ori.collect().to_dicts()

programadas = DatosProgramadas()
asisten = DatosAsisten()
bloqueadas = obtiene_bloqueadas()

### Funciones individuales
def situacion_asisten():
    return (
        pl.read_database(
            query = text('SELECT * FROM asisten'),
            connection = engine,
        )
        .with_columns(
            pl.col('asiste').replace_strict({0: 'No', 1: 'Sí', 2: 'No confirmada'}, return_dtype=pl.Utf8)
        )
        .pivot(
            index='programada_id',
            on='codigo',
            values='asiste',
        )
        .lazy()
        .rename({str(k): v for k, v in universidades.items()})
        .select(['programada_id'] + list(universidades.values()))
    )

def fn_programadas(datos_dic):
    return pl.LazyFrame(datos_dic, schema=schema_programada_op).pipe(str_a_fecha)

def fn_asisten(datos_dic):
    return pl.LazyFrame(datos_dic, schema=schema_asisten)

# datos de las visitas programadas para un determinado mes: muestra todas las visitas, canceladas y no canceladas
def fn_programadas_visita(datos, mes=0):
    programadas = fn_programadas(datos).select(orden_visita)
    if mes != 0:
        programadas = programadas.filter(pl.col('fecha').dt.month() == mes)

    return programadas.sort(['fecha', 'orden']).collect().to_dicts()

# datos de las visitas programadas para una determinada fecha: muestra solo las no canceladas (futuras): Posee argumentos
def fn_programadas_fecha(datos, fecha):
    fecha = fecha_str(fecha)
    return (
        fn_programadas(datos)
        .select(orden_fecha)
        .filter((pl.col('fecha') == fecha) & (pl.col('estatus').is_in(['Confirmada', 'Por confirmar'])))
        .sort(['fecha', 'orden'])
        .with_columns(
            pl.arange(1, pl.len()+1).alias('orden_local'),
        )
        .select(['orden_local', 'rbd', 'nombre', 'organizador', 'estatus'])  # considerar si se necesita 'estatus'
        .rename({'orden_local': 'orden'})
        .collect()
        .to_dicts()
    )

# datos de las visitas organizadas por el usuario: muestra todas las visitas
def fn_programadas_usuario(datos, usuario):
    crt = (pl.col('organizador_id') == usuario) & (pl.col('fecha_lim') >= ahora())
    return (
        fn_programadas(datos)
        .select(orden_modifica)
        .filter(crt)
        .sort(['fecha', 'orden'])
        .with_columns(
            pl.arange(1, pl.len()+1).alias('orden_local'),
            pl.when(pl.col('fecha') >= ahora()).then(pl.lit(0)).otherwise(pl.lit(1)).alias('tipo'),
        )
        .with_columns(
            pl.when((pl.col('tipo') == 0) & (pl.col('estatus') == 'Suspendida')).then(pl.lit(3)).otherwise(pl.col('tipo')).alias('tipo')
        )
        .with_columns(
            pl.when((pl.col('tipo') == 1) & (pl.col('fecha_lim') != fecha_limite)).then(pl.lit(2)).otherwise(pl.col('tipo')).alias('tipo'),
        )
        .select(['programada_id', 'orden_local', 'rbd', 'nombre', 'fecha', 'fecha_lim', 'tipo', 'estatus'])
        .rename({'orden_local': 'orden'})
        .collect()
        .to_dicts()
    )

# datos de las invitaciones recibidas: muestra solo las vigentes no canceladas
def fn_invitaciones(datos, datos_asisten, usuario, condicion):  # tendría que ingresar como parámetro la condición de asistencia de interés (1, 0, 2)
    return (
        fn_asisten(datos_asisten)
        .select(['programada_id', 'codigo', 'asiste', 'invitacion'])
        .filter((pl.col('codigo') == usuario) & (pl.col('asiste') == condicion))
        .join(
            (
                fn_programadas(datos)
                .select(orden_invita)
                .filter((pl.col('fecha') >= ahora()) & (pl.col('estatus').is_in(['Confirmada', 'Por confirmar'])))
            ),
            on='programada_id',
            how='inner'
        )
        .sort(['fecha', 'orden'])
        .with_columns(
            pl.col('comuna_id').replace_strict(comunas, return_dtype=pl.Utf8).alias('comuna'),
            pl.arange(1, pl.len()+1).alias('orden_local'),
        )
        .select(['programada_id', 'orden_local', 'fecha', 'rbd', 'nombre', 'direccion', 'comuna', 'invitacion'])
        .rename({'orden_local': 'orden'})
        .collect()
        .to_dicts()
    )

# asistencia de universidades a determinada visita
def fn_universidades_asisten(datos_asisten, visita):
    return dict(
        fn_asisten(datos_asisten)
        .filter(pl.col('programada_id') == visita)
        .sort(['asiste', 'codigo'])
        .group_by('asiste', maintain_order=True)
        .agg(
            pl.col('codigo')
        )
        .collect()
        .rows()
    )

# asistencia de usuario a determinada visita
def fn_usuario_asiste(datos_asisten, usuario, visita):
    if usuario:
        return (
            fn_asisten(datos_asisten)
            .filter((pl.col('programada_id') == visita) & (pl.col('codigo') == usuario))
            .collect()
            .get_column('asiste')
            .to_numpy()
        )[0]
    else:
        return None

# exporta archivo Excel
def fn_exporta_programada(datos, usuario, mes=0):
    output = io.BytesIO()
    programadas = (
        fn_programadas(datos)
        .with_columns(
            pl.col('comuna_id').replace_strict(comunas, return_dtype=pl.Utf8),
        )
        .select({0: orden_visita}.get(usuario, orden_extendido))
        .rename({0: fn_orden(orden_visita)}.get(usuario, fn_orden(orden_extendido)))
        .sort(['Fecha', 'orden'])
        .drop(['Código', 'orden'])
    )
    if usuario != 0:
        programadas = programadas.join(situacion_asisten(), left_on='ID', right_on='programada_id')
    if mes != 0:
        programadas = programadas.filter(pl.col('Fecha').dt.month() == mes)

    programadas.collect().write_excel(workbook=output, autofilter=False)

    return output.getvalue()

### Interacción con base de datos
Nt_programada = namedtuple('Nt_programada', list(schema_programada.keys())[2:-1], defaults=[None]*21)

dic_reducido = lambda dic: {k: v for k, v in dic.items() if k in list(Nt_programada._fields)}

lista_modifica =  [
    'fecha',
    'direccion',
    'comuna_id',
    'hora_ini',
    'hora_fin',
    'hora_ins',
    'contacto',
    'contacto_tel',
    'contacto_mail',
    'contacto_cargo',
    'orientador',
    'orientador_tel',
    'orientador_mail',
    'asistentes',
    'asistentes_prog',
    'estatus',
    'observaciones',
]

Nt_modifica = namedtuple('Nt_modifica', lista_modifica, defaults=[None]*17)

def ob_programada(tup):
    return Visita(
        organizador_id = tup.organizador_id,
        organizador = tup.organizador,
        fecha = tup.fecha,
        rbd = tup.rbd,
        nombre = tup.nombre,
        direccion = tup.direccion,
        comuna_id = tup.comuna_id,
        hora_ini = tup.hora_ini,
        hora_fin = tup.hora_fin,
        hora_ins = tup.hora_ins,
        contacto = tup.contacto,
        contacto_tel = tup.contacto_tel,
        contacto_mail = tup.contacto_mail,
        contacto_cargo = tup.contacto_cargo,
        orientador = tup.orientador,
        orientador_tel = tup.orientador_tel,
        orientador_mail = tup.orientador_mail,
        asistentes = tup.asistentes,
        asistentes_prog = tup.asistentes_prog,
        estatus = tup.estatus,
        observaciones = tup.observaciones,
    )

def nueva_programada(tup):
    programada = ob_programada(tup)

    with Session(engine) as session:
        session.add(programada)
        session.commit()

def modifica_programada(id_programada, dic):

    with Session(engine) as session:
        modifica = session.query(Visita).filter(Visita.programada_id == id_programada).first()

        for atributo, nuevo_valor in dic.items():
            setattr(modifica, atributo, nuevo_valor)

        session.commit()

def elimina_programada(id_programada):
    with Session(engine) as session:
        elimina = session.query(Visita).filter(Visita.programada_id == id_programada).first()
        session.delete(elimina)

        session.commit()

def modifica_asiste(usuario, id_programada, asiste):
    with Session(engine) as session:
        modifica = session.query(Asiste).filter(Asiste.codigo == usuario).filter(Asiste.programada_id == id_programada).first()
        modifica.asiste = asiste

        session.commit()

def modifica_invitacion(usuario, id_programada, invitacion):
    with Session(engine) as session:
        modifica = session.query(Asiste).filter(Asiste.codigo == usuario).filter(Asiste.programada_id == id_programada).first()
        modifica.invitacion = invitacion

        session.commit()

def cambia_estatus(id_programada, estatus):
    with Session(engine) as session:
        modifica = session.query(Visita).filter(Visita.programada_id == id_programada).first()
        modifica.estatus = estatus
        modifica.fecha_lim = min(modifica.fecha_lim, sig_laboral(ahora(), 1))   # aquí se fija fecha limite.

        session.commit()

### Diseño de la aplicación
linea = html.Hr(style={'borderWidth': '0.15vh', 'width': '100%', 'color': '#104e8b'})
espacio = html.Br()

fecha_sel = max(dia_laboral(), fecha_inicial)
mes_sel = fecha_sel.month

### Barra de navegación
estilo_2c = {'font-size': '36px', 'color': 'white', 'text-align': 'center', 'margin': '0px 30px 0px -45px', 'line-height': '45px'}

mapa_estilo1 = {'size': "sm", 'className': 'componente_inicio'}
mapa_estilo2 = {'size': "sm", 'className': 'componentes'}

navbar = dbc.Navbar(
    dbc.Container(
        dbc.Row(
            [
                # Left column
                dbc.Col(
                    html.Div([
                        dbc.Col(html.Img(src="/assets/cup-logo-3.svg", height="60px", style={'margin-left': '0.5rem'})),
                        dbc.Col(dbc.NavbarBrand(f"{usuarios[usuario]}", id='loc-usuario', className="ms-2", style={'font-size': '14px'})),
                    ]),
                    width=3,
                ),
                # Center column
                dbc.Col([
                    html.P(['Programa de Visitas a Colegios'], className='g-0', style=estilo_2c),
                    html.P(['2026'], className='g-0', style=estilo_2c),
                    ], className="justify-content-center",
                    width=5,
                    style={'marginLeft': 0}
                ),
                # Right column
                dbc.Col(
                    dbc.Nav(
                        [
                            dbc.NavItem(dbc.Button("Visitas", id='btn-resumen', color="#175a96", size="sm", className='componente-inicio')),
                            dbc.NavItem(dbc.Button("Agregar", id='btn-agrega', color="#175a96", size="sm", className='componentes')),
                            dbc.NavItem(dbc.Button("Modificar/Eliminar", id='btn-elimina', color="#175a96", size="sm", className='componentes')),
                            dbc.NavItem(dbc.Button("Invitaciones", id='btn-invita', color="#175a96", size="sm", className='componentes')),
                        ],
                        className="ms-auto",
                        navbar=True,
                    ),
                    className="justify-content-center",
                    width=4,
                ),
            ],
            className="g-0 w-100",
            align="center",
        ),
        fluid=True,
    ),
    color="#175a96",
    dark=True,
)

### Acceso
acceso = html.Div(
    dbc.Container([
        dbc.Row(
            dbc.Col(
                html.H2(['Ingreso de Usuario'], style={'textAlign': 'center', 'color': '#7F7F7F', 'margin': 0}),
            )
        ),
        dbc.Row(
            html.Div([
                dbc.Row([
                    dbc.Col(html.P(html.Span('Usuario:'), className='pAcceso'), width=2),
                    dbc.Col(
                        dcc.Dropdown(opciones(universidades), placeholder='Seleccione su universidad', id='sel-u', style={'width': '100%', 'textAling': 'left', 'background': '#f1f1f1', 'border': '0px'}),
                        align='center',
                        width=10,
                    ),
                ]),
                dbc.Row([
                    dbc.Col(html.P(html.Span('Password:'), className='pAcceso'), width=2),                
                    dbc.Col(dbc.Input(placeholder="Ingrese su password", id='inp-pw', type="password", debounce=True, style={'background': '#f1f1f1', 'border': '0px', 'borderRadius': '0px'}), 
                        align='center',
                        width=6,
                    ),
                ])
            ], className='boxAcceso',
            ), justify='center',
        ),
        dbc.Row([
                html.Button('Ingresar', id='btn-ingresar', n_clicks=0, className='btn btn-outline-primary mx-1', style={'width': '20%'}),
                html.Button('Cancelar', id='btn-cancelar', n_clicks=0, className='btn btn-outline-primary mx-1', style={'width': '20%'}),
            ],
            justify='center',
        )
    ],
    style={'marginTop': '5%', 'marginBottom': '5%'}
    ),
    id='acceso',
)

modal_acceso = dbc.Modal(
    [
        dbc.ModalBody(acceso),
    ],
    id="modal-acceso",
    is_open=False,
    size='lg',
)

### Advertencias
def form_advertencia(texto):
    return html.Div([
        dbc.Row([
            dbc.Col(
                html.P([texto], id='txt-advert', style={'color': 'white', 'marginBottom': '0px', 'padding': '6px 12px'})
            )
        ], align='center')
    ],
    style={'background': '#c61c22'}
)

### Inicio
op_meses = opciones_meses()

def botones_mes(mes):
    return html.Div(
        dbc.Row([
            dbc.Col(html.H5('Filtro según mes:', style={'marginBottom': 0}), width='auto'),
            dbc.Col(
                dcc.RadioItems(
                    id = 'selec-mes',
                    options = op_meses,
                    value = mes,
                    inline = True,
                    labelStyle = {'display': 'inline-block', 'fontSize': '14px', 'fontWeight': 'normal'},
                    inputStyle = {'marginRight': '5px', 'marginLeft': '20px'},
                ),
                width='auto'
            ),
            dbc.Col(html.Button("Calendario", id='abre-calendario',  className="btn btn-outline-primary",
                               style={'width': '103%', 'margin': '10px 0px 10px 15px', 'padding': '3px 15px'}),
                    width='auto', className='ms-auto', style={'marginRight': '20px'}),
        ], align='center')
    )

getRowStyle = {
    'styleConditions': [{
        'condition': 'params.rowIndex % 2 === 0',
        'style': {'backgroundColor': 'rgb(47, 164, 231, 0.1)'},
    }],
}

columnDefs_programadas = [
    {'field': 'fecha', 'cellStyle': {'textAlign': 'center'}},
    {'field': 'rbd', 'headerName': 'RBD', 'cellStyle': {'textAlign': 'center'}, 'filter': True},
    {'field': 'nombre', 'width': 450, 'filter': True},
    {'field': 'organizador', 'width': 450, 'filter': True, 'sortable': True},
    {'field': 'estatus', 'filter': True, 'sortable': True},
]

def grid_programadas(datos):
    return dag.AgGrid(
        id='grid-programadas',
        rowData=datos,
        defaultColDef={'resizable': True},
        columnDefs=columnDefs_programadas,
        dashGridOptions = {
            'rowSelection': 'single',
        },
        columnSize='sizeToFit',
        getRowStyle=getRowStyle,
        style={'height': '65vh', 'width': '100%'}
    )

btn_exp_visitas = dbc.Row([
    html.Button('Exportar a Excel', id='exporta-visitas', className='btn btn-outline-primary',
                style={'width': '15%', 'margin': '10px 15px', 'padding': '6px 20px'}),
    dcc.Download(id='exporta-visitas-archivo'),
], justify='end',)

#### Reporte visita
fto_espanol = lambda x: convierte_esp(f'{datetime.strptime(x, "%Y-%m-%d").date():%A, %d %B %Y}')
fto_blanco = lambda x: '' if x == None else x
fto_hora = lambda x: x[:-3] + ' hrs.' if x != '00:00:00' else ''
fto_hora = lambda x: x[:-3] + ' hrs.' if x != None else ''
fto_hora = lambda x: horas_15[x] + ' hrs.' if x != None else None

formato_items = {
    'fecha': fto_espanol,
    'direccion': fto_blanco,
    'comuna_id': lambda x: comunas[x],
    'hora_ins': fto_hora,
    'hora_ini': fto_hora,
    'hora_fin': fto_hora,
    'orientador': fto_blanco,
    'asistentes': fto_blanco,
}

def da_formato(datos_dic):
    return {k: formato_items.get(k, lambda x: x)(datos_dic[k]) for k in datos_dic.keys()}

items_reporte = ['estatus','organizador', 'nombre', 'rbd', 'fecha', 'direccion', 'comuna_id', 'hora_ins', 'hora_ini', 'hora_fin', 'orientador', 'asistentes']
items_reporte_label = ['Estatus', 'Organizador', 'Nombre', 'RBD', 'Fecha', 'Dirección', 'Comuna', 'Hora instalación', 'Hora inicio', 'Hora término', 'Orientador/a', 'Asistentes']
items_reporte_dic = dict(zip(items_reporte, items_reporte_label))

def info_gral(item, datos_dic):
    return html.Div([
        html.P(items_reporte_dic[item], style={'width': '18%', 'fontSize': '16px', 'marginLeft': 20, 'marginBottom': -2, 'display': 'inline-block'}),
        html.P(': ' + str(datos_dic[item]), style={'width': '75%', 'fontSize': '16px', 'marginBottom': -2, 'display': 'inline-block'}),
    ])

def seccion_info_gral(datos_dic={}):
    if datos_dic == {}:
        return None
    else:
        dic_reducido = {k: datos_dic.get(k, None) for k in items_reporte}
        dic_reducido_fto = da_formato(dic_reducido)
        return html.Div(
            [html.H6('Información general:', style={'fontSize': '17px'})] +
            [info_gral(item, dic_reducido_fto) for item in items_reporte]
        )

def universidad_asiste(n, universidad):
    return html.Div(
        html.P(str(n) + ') ' + universidades[universidad], style={'marginLeft': 20, 'marginBottom': -2, 'fontSize': '16px', 'fontWeight': 'normal'})
    )

def seccion_universidades_asisten(dic, crt):
    lista = dic.get(crt, None)
    dic_opcion = {0: 'Universidades que no asisten:', 1: 'Universidades que asisten:', 2: 'Universidades sin confirmación:'}
    if lista != None:
        return html.Div([
            html.H6(dic_opcion[crt], style={'fontSize': '17px', 'marginBottom': '4px'}),
            dbc.Col(
                [universidad_asiste(n, univ) for n, univ in enumerate(lista, start=1)]
            )
        ], style={'marginBottom': '15px'})

def universidades_asisten_criterio(lista, crt):
    return html.Div([
        html.H6(crt, style={'fontSize': '17px', 'marginBottom': '4px'}),
        dbc.Col(
            [universidad_asiste(n, univ) for n, univ in enumerate(lista, start=1)]
        )
    ], style={'marginBottom': '15px'})
    

def seccion_universidades_asisten2(dic):
    dic_opcion = {0: 'Universidades que no asisten:', 1: 'Universidades que asisten:', 2: 'Universidades sin confirmación:'}
    lista_retorno = []
    for i in [1, 0, 2,]:
        lista = dic.get(i, None)
        crt = dic_opcion[i]
        if lista != None:
            lista_retorno.append(universidades_asisten_criterio(lista, crt))
    return lista_retorno

def seccion_asistencia(datos_dic):
    valor = '' if datos_dic['asistentes_prog'] == None else datos_dic['asistentes_prog']
    return html.Div(
        dbc.Row(
            [
                dbc.Col(html.P(['Asistencia/', html.Br(), 'Programación'], style={'marginLeft': 20, 'lineHeight': '1.1'}), width=3),
                dbc.Col(dcc.Textarea(value=valor, style={'borderColor': '#c7c7c7', 'marginLeft': '-40px', 'width': '100%', 'color': '#5c5c5c'}), width=9),
            ],
            align="start",
        ),
        style={'marginTop': 10},
    )
    
def seccion_observaciones(datos_dic):
    valor = '' if datos_dic['observaciones'] == None else datos_dic['observaciones']
    return html.Div(
        dbc.Row(
            [
                dbc.Col(html.P('Observaciones', style={'marginLeft': 20}), width=3),
                dbc.Col(dcc.Textarea(value=valor, style={'borderColor': '#c7c7c7', 'marginLeft': '-40px', 'width': '100%', 'color': '#5c5c5c'}), width=9),
            ],
            align="start",
        ),
        style={'marginTop': 10}
    )

opciones_asiste = [
   {'label': 'Sí', 'value': 1},
   {'label': 'No', 'value': 0},
   {'label': 'No confirmada', 'value': 2, 'disabled': True},
]

opciones_asiste2 = [
   {'label': 'Sí', 'value': 1, 'disabled': True},
   {'label': 'No', 'value': 0, 'disabled': True},
   {'label': 'No confirmada', 'value': 2, 'disabled': True},
]

def fn_opciones_asiste(fecha):
    fecha = fecha_str(fecha)
    if fecha >= ahora():
        return opciones_asiste
    else:
        return opciones_asiste2

def selector_asiste(usuario, valor, fecha):
    return html.Div([
        linea,
        html.H6('Asistencia a visita:', style={'fontSize': '17px', 'margin': 0}),
        dbc.Row([
            dbc.Col([
                html.P(f'{universidades[usuario]}:', style={'fontSize': '16px', 'margin': '2px 0px 0px 20px'}),
            ], width='auto', style={'padding': 0}),
            dbc.Col([
                dcc.RadioItems(
                    id = 'selector-asiste',
                    options=fn_opciones_asiste(fecha),
                    value = valor,
                    inline = True,
                    labelStyle = {'display': 'inline-block', 'fontSize': '16px', 'fontWeight': 'normal'},
                    inputStyle = {'marginRight': '5px', 'marginLeft': '20px'},
                )
            ], width='auto', style={'padding': 0, 'marginTop': '2px'}),
            dbc.Col(
                dbc.Button('Aplicar cambio', id='btn-selector-asiste', outline=True, color="primary", className='me-2', disabled=True,
                           style={'fontSize': '14px', 'padding': '4px 28px', 'marginTop': '6px', 'marginLeft': '8px'})
            )
        ]),
    ], style={'marginLeft': 10, 'marginBottom': 1, 'marginTop': 20})

reporte_programada = html.Div(
    dbc.Modal(
        [
            dbc.ModalHeader(html.H4('Información de la visita', style={'fontSize': '2rem', 'marginBottom': '0px'}), close_button=False),
            dbc.ModalBody(id='reporte-prog-contenido'),
            dbc.ModalFooter(
                html.Div([
                    html.Div(
                        [dbc.Button('Descargar reporte', id='descarga-reporte', outline=True, color="primary", className='me-2', style={'padding': '6px 25px'})],
                        id='div-descarga-reporte',
                        hidden=True,
                        style={'display': 'inline-block'},
                    ),
                    dbc.Button('Cerrar', id='btn-cerrar-reporte-prog', outline=True, color="primary", className='me-2', style={'padding': '6px 25px'}),
                    dcc.Download(id='descarga-reporte-archivo'),
                ])
            ),
        ],
        id='modal-reporte-prog',
        size='lg',
        keyboard=False,
        backdrop="static",
   ),
)

def reporte_reducido(datos, asisten):
    return html.Div([
        seccion_info_gral(datos),
        linea,
        seccion_universidades_asisten(asisten, 1),
    ])

def reporte_extendido(datos, asisten, usuario, uasiste, fecha):
    return html.Div([
        seccion_info_gral(datos),
        seccion_asistencia(datos),
        seccion_observaciones(datos),
        linea,
        html.Div(seccion_universidades_asisten(asisten, 1), id='rep-asiste'),
        html.Div(seccion_universidades_asisten(asisten, 0), id='rep-no-asiste'),
        html.Div(seccion_universidades_asisten(asisten, 2), id='rep-no-confirma'),
        selector_asiste(usuario, uasiste, fecha),
    ])

def escoge_reporte(usuario, datos, asisten, uasiste, fecha):
    if usuario:
        return reporte_extendido(datos, asisten, usuario, uasiste, fecha)
    else:
        return reporte_reducido(datos, asisten)

#### Reporte pdf
sello = './assets/sello4.png'
cup = './assets/cup-logo3.png'

def exporta_reporte(visita, asisten):
    output = io.BytesIO()

    cm = 72. / 2.54
    top = 792. - (1.3 * cm) - 22
    step = 18
    canvas = Canvas(output, pagesize=(612., 792.))

    canvas.setFont('Helvetica', 22)
    canvas.setFillColor(HexColor('#1b81e5'))
    canvas.drawCentredString(306 + cm, top, 'Programa de Visitas a Colegios 2026')
    top -= (step + 14)

    canvas.setFont('Helvetica', 14)
    canvas.drawString((2 * cm), top, 'Antecedentes de la Visita')
    top -= (step + 3)

    canvas.setFont('Helvetica', 11)
    canvas.setFillColor(HexColor('#596a6d'))
    for item in items_reporte:
        canvas.drawString((2.3 * cm), top, f'{items_reporte_dic[item]}')
        canvas.drawString((2.3 * cm)*2.22, top, ':')
        canvas.drawString((2.3 * cm)*2.36, top, f'{formato_items.get(item, lambda x: x)(visita[item])}')
        top -= step

    top -= (step - 2)
    canvas.setFont('Helvetica', 14)
    canvas.setFillColor(HexColor('#1b81e5'))
    canvas.drawString((2. * cm), top, 'Universidades Participantes')
    top -= (step + 3)

    canvas.setFont('Helvetica', 11)
    canvas.setFillColor(HexColor('#596a6d'))
    for asiste in asisten:
        canvas.drawString((2.3 * cm), top, f"{universidades[asiste]}")
        top -= step

    canvas.drawImage(cup, 1.8*cm, 644, width=80, height=None, mask='auto', preserveAspectRatio=True)

    # sección timbre:
    # sello:
    canvas.setFillAlpha(0.5)
    canvas.drawImage(sello, 418, -150, width=125, height=None, mask='auto', preserveAspectRatio=True)
    canvas.setFillAlpha(1)
    # línea:
    canvas.line(430, 1.9*cm, 530, 1.9*cm)
    # texto:
    canvas.drawCentredString(480, 1.5*cm, 'TIMBRE')

    canvas.save()

    retorna = output.getvalue()
    output.close()

    return retorna

#### Calendario
with open('./data/indice_calendario.pkl', 'rb') as f:
    idx_calendario = pickle.load(f)

with open('./data/shape_calendario.pkl', 'rb') as f:
    shape_calendario = pickle.load(f)

botones_mes_cal = dbc.Row([
    dbc.Col(html.H5('Selección de mes:', style={'marginBottom': 0}), width='auto'),
    dbc.Col(
        dcc.RadioItems(
            id = 'selec-mes-cal',
            options = op_meses[1:],
            value = None,
            inline = True,
            labelStyle = {'display': 'inline-block', 'fontSize': '14px', 'fontWeight': 'normal'},
            inputStyle = {'marginRight': '5px', 'marginLeft': '20px'},
        ),
        width='auto'
    ),
    dbc.Col(html.Button("Imprimir", id='imprimir-calendario',  className="btn btn-outline-primary",
                       style={'width': '100%', 'margin': '10px 0px 10px 15px', 'padding': '3px 15px'}),
            width='auto', className='ms-auto', style={'marginRight': '20px'}),
],
align='center'
)

semana = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes']
meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']

def dic_eventos(base, mes):
    return dict(
        base.select(['orden', 'fecha', 'nombre'])
        .filter(pl.col('fecha').dt.month() == mes)
        .sort(['fecha', 'orden'])
        .drop('orden')
        .with_columns(
            pl.col('fecha').dt.day(),
            pl.col('nombre').str.slice(0, 40)
        )
        .group_by('fecha', maintain_order=True)
        .agg(
            pl.col('nombre')
        )
        .rows()
    )

def agrega_evento(mes, dia, evento, eventos):
    idx_s, idx_d = idx_calendario[mes][dia]
    eventos[idx_s][idx_d] = evento

def crea_figura(base, mes):

    isem, idia = shape_calendario[mes]
    eventos = [[[] for i in range(idia)] for j in range(isem)]

    dic = dic_eventos(base, mes)
    for dia, evento in dic.items():
        agrega_evento(mes, dia, evento, eventos)

    plot_defaults = dict(
        figsize=(18, 1.1*isem),
        sharex=True,
        sharey=True,
        dpi=80,
    )

    fig, axs = plt.subplots(isem, 5, **plot_defaults)
    fig.tight_layout(rect=(0, 0, 1, 0.925))

    for dia, pos in idx_calendario[mes].items():
        ax = axs[pos[0], pos[1]]
        ax.set_xticks([])
        ax.set_yticks([])
        ax.text(.02, .9, str(dia), verticalalignment='top', horizontalalignment='left', fontsize=14)

        visitas = "\n".join(eventos[pos[0]][pos[1]])
        ax.text(.03, .65, visitas, verticalalignment='top', horizontalalignment='left', fontsize=9)

    for n, day in enumerate(semana):
        axs[0][n].set_title(day, fontsize=14, loc='center')

    fig.subplots_adjust(hspace=0, wspace=0)
    fig.suptitle(f'{meses[mes-1]} 2026', fontsize=20, x=0.503, y=1)

    output = BytesIO()
    fig.savefig(output, format='png', bbox_inches='tight')

    fig_data = base64.b64encode(output.getbuffer()).decode('ascii')  # 'utf-8'
    plt.close()

    return f'data:image/png;base64,{fig_data}'

calendario = html.Div(
    dbc.Modal(
        [
            dbc.ModalHeader(html.H4('Calendario de visitas', style={'fontSize': '2rem', 'marginBottom': '0px'}), close_button=False),
            dbc.ModalBody([
                html.Img(id='calendario-contenido',
                    style={
                        'max-width': '100%',
                        'height': 'auto',
                        'display': 'block',
                        'margin-left': 'auto',
                        'margin-right': 'auto',
                    }
                ),
                dbc.Row([
                    dbc.Col(html.H5('Selección del mes:', style={'fontSize': '14px', 'marginBottom': 0}), width='auto', style={'padding': '2px 4px 0px'}),
                    dbc.Col(
                        dcc.RadioItems(
                            id = 'selec-mes-cal',
                            options = op_meses[1:],
                            value = None,
                            inline = True,
                            labelStyle = {'display': 'inline-block', 'fontSize': '12px', 'fontWeight': 'normal'},
                            inputStyle = {'marginRight': '5px', 'marginLeft': '17px'},
                        ),
                        width='auto'
                    ),
                    dbc.Col([dbc.Button("Imprimir", id='imprimir-calendario', outline=True, color="primary", className='me-2',
                                       style={'width': '115%', 'fontSize': '13px', 'padding': '4px 20px'}),
                            dcc.Download('exporta-calendario')],
                            width='auto', className='ms-auto', style={'marginRight': '8px'}),
                ], align='center', style={'margin': '4px 0px 0px'}),
            ], style={'paddingBottom': '10px'}),
            dbc.ModalFooter(
                html.Div([
                    dbc.Button('Cerrar', id='btn-cerrar-calendario', outline=True, color="primary", className='me-2', style={'padding': '6px 25px'}),
                ]),
                style={'paddingTop': '6px'}
            ),
        ],
        id='modal-calendario',
        size='xl',
        keyboard=False,
        backdrop="static",
    )
)

#### Forma
def form_inicio(datos, mes):
    return dbc.Form([
        html.H3(['Visitas Programadas'], className='titulo_pagina'),
        linea,
        html.Div(botones_mes(mes)),
        html.Div(grid_programadas(datos)),
        html.Div(reporte_programada),
        html.Div(calendario),
        html.Div(btn_exp_visitas),
    ])


### Agrega
def disabled_btn(fecha):
    if isinstance(fecha, str):
        fecha = datetime.strptime(fecha, '%Y-%m-%d').date()
    return (fecha in bloqueadas)

op_horas = opciones(horas_15)
op_colegios = opciones(colegios)
op_comunas = opciones(comunas)

# selección de fecha
columnDefs_ing = [
    {'field': 'orden', 'headerName': 'N', 'width': 70, 'cellStyle': {'textAlign': 'center'}},
    {'field': 'rbd', 'headerName': 'RBD', 'width': 150, 'cellStyle': {'textAlign': 'center'}, 'filter': True},
    {'field': 'nombre', 'width': 450, 'filter': True},
    {'field': 'organizador', 'width': 450, 'filter': True, 'sortable': True},
    {'field': 'estatus', 'filter': True, 'sortable': True},
]

def fecha_visita(dia, datos):
    return html.Div([
        dbc.Row([
            dbc.Col([
                html.H5('Seleccione una fecha:'),
                dcc.DatePickerSingle(
                    id='pick-fecha',
                    min_date_allowed=dia_laboral(),
                    max_date_allowed=fecha_final,
                    disabled_days=feriados,
                    first_day_of_week=1,
                    initial_visible_month=str(mes_sel),
                    date=dia,
                    display_format='D MMM YYYY',
                    stay_open_on_select=False,
                    show_outside_days=False,
                )
            ],
            style={'width': '18%'}
            ),
            dbc.Col([
                html.H5('Visitas programadas para dicha fecha'),
                dag.AgGrid(
                    id='grid-agrega',
                    rowData=datos,
                    defaultColDef={'resizable': True},
                    columnDefs=columnDefs_ing,
                    columnSize='sizeToFit',
                    dashGridOptions = {
                        'domLayout': 'autoHeight',
                        'rowSelection': 'single',
                    },
                    style={'width': '100%'},
                ),
            ],
            style={'width': '82%'},
            ),
        ]),
        linea,
    ])

# selector de colegio
texto_colegio = '(Seleccione directamente de la lista o ingrese alguna palabra del nombre para reducir las opciones)'

def colegio():
    return html.Div([
        dbc.Row([
            dbc.Col(
                html.H5('Seleccione un colegio: '), style={'width': '18%'}
            ),
            dbc.Col([
                html.H5('Según RBD'),
                dbc.Input(type='number', id='sel-rbd', placeholder='Ingrese RBD', required=True, debounce=True, style={'width': '100%', 'padding': '5px 12px'}),
                dbc.FormText('(Sin dígito verificador)', style={'font-size': '12px', 'marginTop': '0px'}),
            ], style={'width': '15%'}),
            dbc.Col([
                html.H5('Según nombre'),
                dcc.Dropdown(op_colegios, id='sel-nombre', placeholder='Seleccione de la lista / Ingrese palabra(s) del nombre'),
                dbc.FormText(texto_colegio, style={'font-size': '12px', 'marginTop': '0px'}),
            ], style={'width': '65%'}),
        ]),
        dbc.Row(
            dbc.Col(
                dbc.Button('Limpiar selección', id='btn-limpiar-sel', outline=True, color="primary", className='me-2', n_clicks=0,
                           style={'font-size': '14px', 'width': '15%', 'padding': '3px 12px'}),
            )
        ),
        linea,
    ])

# dirección del colegio
def direccion():
    return html.Div([
        dbc.Row([
            dbc.Col(
                html.H5(['Dirección del', html.Br(), 'establecimiento:']),
                style={'width': '18%'}),
            dbc.Col([
                html.H5(['Dirección']),
                dbc.Input(type='text', id='id-direccion', placeholder='Dirección (máx. 200 caracteres)', debounce=True),
            ], style={'width': '50%'}),
            dbc.Col([
                html.H5(['Comuna']),
                dcc.Dropdown(op_comunas, id='id-comuna'),
            ], style={'width': '18%'}),
        ]),
        linea,
    ])

def horario():
    return html.Div([
        dbc.Row([
            dbc.Col(
                html.H5(['Especifique el horario:']),
                style={'width': '18%'}
            ),
            dbc.Col([
                html.H5(['De inicio']),
                dcc.Dropdown(op_horas, id='hr-inicio', style={'margin-right': '45px'})  #, persistence=True, persistence_type='memory')
            ], style={'width': '15%'}),
            dbc.Col([
                html.H5(['De término']),
                dcc.Dropdown(op_horas, id='hr-termino', style={'margin-right': '45px'})  #, persistence=True, persistence_type='memory')
            ], style={'width': '15%'}),
            dbc.Col([
                html.H5(['De instalación']),
                dcc.Dropdown(op_horas, id='hr-instala', style={'margin-right': '45px'})  #, persistence=True, persistence_type='memory')
            ], style={'width': '15%'}),
        ]),
        linea,
    ])

def contacto():
    return html.Div([
        dbc.Row([
            dbc.Col(
                html.H5(['Añada información', html.Br(), 'de contacto:']),
                style={'width': '18%'}
            ),
            dbc.Col([
                html.H5(['Contacto']),
                dbc.Input(type='text', id='contacto-nom', placeholder='Nombre (máx. 60 caracteres)', debounce=True),
                dbc.Input(type='text', id='contacto-cel', placeholder='Celular (máx. 20 caracteres)', debounce=True),
                dbc.Input(type='text', id='contacto-mail', placeholder='Correo (máx. 50 caracteres)', debounce=True),
                dbc.Input(type='text', id='contacto-cargo', placeholder='Cargo (máx. 40 caracteres)', debounce=True),
            ], style={'width': '40%'}),
            dbc.Col([
                html.H5(['Orientador']),
                dbc.Input(type='text', id='orienta-nom', placeholder='Nombre (máx. 60 caracteres)', debounce=True),
                dbc.Input(type='text', id='orienta-cel', placeholder='Celular (máx. 20 caracteres)', debounce=True),
                dbc.Input(type='text', id='orienta-mail', placeholder='Correo (máx. 50 caracteres)', debounce=True),
            ], style={'width': '40%'}),
        ]),
        linea,
    ])

# asistencia
def asistencia():
    return html.Div([
        dbc.Row([
            dbc.Col(
                html.H5('Asistencia estimada:'),
                style={'width': '18%'}
            ),
            dbc.Col([
                html.H5('Cantidad'),
                dbc.Input(type='text', id='cantidad', placeholder='Cantidad de alumnos/as (máx. 80 caracteres)', debounce=True),
            ], style={'width': '30%'}),
            dbc.Col([
                html.H5('Detalle asistencia/programa'),
                dbc.Textarea(placeholder='Detalle de asistencia y programa', id='cant-detalle',
                             style={'width': '97%', 'height': '4.5rem', 'border': '1px solid #d5d5d5', 'borderRadius': '5px'}),
            ], style={'width': '52%'}),
        ]),
        linea,
    ])

lista_estatus = ['Confirmada', 'Por confirmar']

def estatus():
    return html.Div([
        dbc.Row([
            dbc.Col([
                html.H5(['Indique el estatus:']),
                dcc.Dropdown(lista_estatus, id='def-estatus', style={'width': '95%'})
            ], style={'width': '18%'}),
            dbc.Col([
                html.H5(['Observaciones:']),
                dbc.Textarea(id='txt-observa', placeholder='Observaciones pertinentes de distinta índole', 
                             style={'width': '100%', 'height': '6rem', 'border': '1px solid #d5d5d5', 'borderRadius': '5px'}),
            ], style={'width': '52%'}),
        ]),
        linea,
    ])

def botones_agrega():
    return html.Div(
        dbc.Row([
            dbc.Col([
                dbc.Button('Agregar visita', id='btn-agregar-visita', outline=True, color="primary", className='me-2', n_clicks=0, disabled=disabled_btn(dia_laboral()),
                           style={'width': '15%'}),
                dbc.Button('Limpiar campos', id='btn-limpiar', outline=True, color="primary", className='me-2', n_clicks=0, style={'width': '15%'}),
            ])
        ], style={'margin-bottom': '2rem'})
    )

# modal que informa que fecha no está disponible: agrega
fecha_no_disponible = html.Div(
    dbc.Modal(
        [
            dbc.ModalHeader(html.H4('No es posible agregar visita'), close_button=False),
            dbc.ModalBody(html.Div('La fecha escogida ya no está disponible para agregar una nueva visita. Algún otro usuario la ocupó en el intertanto.')),
            dbc.ModalFooter(dbc.Button('Cerrar', id='cierra-fecha-no-disponible', outline=True, color="primary", className='me-2', n_clicks=0, style={'width': '20%'})),
        ],
        id='modal-fecha-no-disponible',
        size='lg',
        centered=True,
    ),
)

#### Forma
def form_agrega(dia, datos):
    return html.Div([
        html.Div(html.H3('Agrega Nueva Visita', className='titulo_pagina')),
        linea,
        fecha_visita(dia, datos),
        colegio(),
        direccion(),
        horario(),
        contacto(),
        asistencia(),
        estatus(),
        botones_agrega(),
        fecha_no_disponible,
    ], id='forma-agrega')

### Modifica/Elimina
getRowStyle_modifica = {
    "styleConditions": [
        {
            "condition": "params.data.tipo == 1",  # vencida no confirmada realización: amarillo
            "style": {"backgroundColor": "#FBFF9E"},
        },
        {
            "condition": "params.data.tipo == 2",  # vencida confirmada realización: verde
            "style": {"backgroundColor": "#B6FFA4"},
        },
        {
            "condition": "params.data.tipo == 3",  # no vencida suspendida: rojo
            "style": {"backgroundColor": "#FFD1D1"},
        },
    ],
    "defaultStyle": {"backgroundColor": "#fff"},  # no vencida: blanco
}

columnDefs_modifica = [
    {'field': 'orden', 'headerName': 'N', 'width': 60, 'cellStyle': {'textAlign': 'center'}},
    {'field': 'rbd', 'headerName': 'RBD', 'width': 120, 'cellStyle': {'textAlign': 'center'}, 'filter': True},
    {'field': 'nombre', 'width': 510, 'filter': True},
    {'field': 'fecha', 'cellStyle': {'textAlign': 'center'}},
]

def grid_modifica(datos):
    return dag.AgGrid(
        id='grid-modifica',
        rowData=datos,
        defaultColDef={'resizable': True},
        columnDefs=columnDefs_modifica,
        columnSize='sizeToFit',
        dashGridOptions = {
            'domLayout': 'autoHeight',
            'rowSelection': 'single',
        },
        getRowStyle=getRowStyle_modifica,
        style={'width': '95%', 'margin': 0},
    )

botones_modifica = [
    dbc.Button('Modificar', id='btn-mod-visita', outline=True, color="primary", className='me-2', n_clicks=0, disabled=True, style={'width': '66%', 'marginBottom': '4px'}),
    dbc.Button('Eliminar', id='btn-elim-visita', outline=True, color="primary", className='me-2', n_clicks=0, disabled=True, style={'width': '66%', 'marginTop': '4px'}),
]

estilo_iconos = {'size': 15, 'margin': '0px 20px 0px 10px', 'display': 'inline-block'}
iconos_modifica = html.Div([
    html.Img(src='./assets/blanco3.png', height='25px'),
    html.P('Por realizarse', style=estilo_iconos),
    html.Img(src='./assets/rojo3.png', height='25px'),
    html.P('Suspendida', style=estilo_iconos),
    html.Img(src='./assets/amarillo3.png', height='25px'),
    html.P('De fecha anterior, no confirmada', style=estilo_iconos),
    html.Img(src='./assets/verde3.png', height='25px'),
    html.P('De fecha anterior, confirmada', style=estilo_iconos),
], style={'margin': '10px 0px'})

nota_modifica = html.Div([
    html.P('Nota:', style={'marginBottom': '0px'}),
    html.P([
        'Las visitas cuya fecha de realización no se ha alcanzado aparecen generalmente en blanco, salvo que haya sido "Suspendida", en cuyo caso', html.Br(),
        'aparece en rojo. Alcanzada la fecha de realización de una visita, ésta cambia de color a amarillo indicando que se debe confirmar si fue realizada', html.Br(),
        'o no. Para realizar esta confirmación seleccione una determinada visita y escoja en la ventana emergente la opción que corresponda ("Ralizada",', html.Br(),
        '"No realizada"). En ese momento la visita cambiará de color a verde y permanecerá así por tres día habiles, permitiendo corregir cualquier error.', html.Br(),
        'Si no se realiza esta acción la visita permanecerá indefinidamente como no confirmada.'
    ], style={'fontSize': '14px', 'marginBottom': '30px'})
])

# modal para confirmar realización
confirma_realizacion = html.Div(
    dbc.Modal(
        [
            dbc.ModalHeader(html.H4('Realización de la visita', style={'fontSize': '1.5rem', 'marginBottom': '0px'}), close_button=False),
            dbc.ModalBody(
                dbc.Row([
                    dbc.Col([
                        html.P('Confirme si la visita fue realizada o no.', style={'fontSize': '18px', 'margin': '0px 21px'}),
                        dcc.RadioItems(
                            id = 'mod-selector-asiste',
                            options=[
                               {'label': 'Realizada', 'value': 'Realizada'},
                               {'label': 'No realizada', 'value': 'Suspendida'},
                            ],
                            inline = True,
                            labelStyle = {'display': 'inline-block', 'fontSize': '16px', 'fontWeight': 'normal'},
                            inputStyle = {'marginRight': '5px', 'marginLeft': '20px'},
                        )
                    ], width='auto', style={'padding': 0, 'margin': '2px, 40px'}),
                ]),
                id='confirma-visita'
            ),
            dbc.ModalFooter(
                html.Div([
                    dbc.Button('Cerrar', id='btn-cerrar-confirma-visita', outline=True, color="primary", className='me-2', style={'fontSize': '15px', 'padding': '4px 25px'}),
                    dbc.Button('Aplicar cambio', id='cambia-confirma-visita', outline=True, color="primary", className='me-2', style={'fontSize': '15px', 'padding': '4px 25px'}),
                ])
            ),
        ],
        id='modal-confirma-visita',
        keyboard=False,
        backdrop="static",
   ),
)

# modal para confirmar eliminación de visita
confirma_eliminacion = html.Div(
    dbc.Modal(
        [
            dbc.ModalHeader(html.H4('Eliminación de visita', style={'fontSize': '1.5rem', 'marginBottom': '0px'}), close_button=False),
            dbc.ModalBody(
                dbc.Row([
                    dbc.Col([
                        html.P('¿Desea eliminar permanentemente esta visita?', style={'fontSize': '18px', 'margin': '0px 20px'}),
                        html.Div([
                            dbc.Button('Sí', id='btn-confirma-elimina', n_clicks=0, outline=True, color="primary", className='me-2', style={'fontSize': '15px', 'padding': '4px 25px'}),
                            dbc.Button('No', id='btn-confirma-no-elimina', n_clicks=0, outline=True, color="primary", className='me-2', style={'fontSize': '15px', 'padding': '4px 25px'}),
                        ], style={'margin': '5px 20px'})
                    ], width='auto', style={'padding': 0, 'margin': '2px, 40px'}),
                ]),
                id='confirma-elimina'
            ),
        ],
        id='modal-confirma-elimina',
        keyboard=False,
        backdrop="static",
   ),
)

#### Modifica
nota_pie = lambda x: '()' if x == None else f'({x})'

lista_estatus_modifica = ['Confirmada', 'Por confirmar', 'Suspendida']

# colegio: esta información no se modifica
def identificacion(nombre, rbd):
    return html.Div([
        dbc.Row([
            dbc.Col([
                html.Div(
                    html.P(f'Colegio: {nombre}\nRBD: {rbd}', style={'fontSize': '20px', 'margin': '0px'}),
                    style={'whiteSpace': 'pre-line', 'background': '#f7f7f7', 'padding': '12px 25px'}, 
                )
            ])
        ]),
        linea,
    ])

# estatus
def mod_estatus(estatus):
    return html.Div([
        dbc.Row([
            dbc.Col(html.H5(['Estatus:'], style={'width': 'auto', 'marginTop': 10}), width=1),
            dbc.Col(dcc.Dropdown(lista_estatus_modifica, value=estatus, id='mod-def-estatus'), width=2),
            dbc.Col(html.P(nota_pie(estatus), style={'margin': '4px 0px', 'color': 'red'}), width=2),
        ], align='center'),
        linea,
    ])

# dirección del colegio
def mod_direccion(direccion, comuna):
    return html.Div([
        dbc.Row([
            dbc.Col(
                html.H5(['Dirección del', html.Br(), 'establecimiento:']),
                style={'width': '18%'}),
            dbc.Col([
                html.H5(['Dirección']),
                dbc.Input(type='text', id='mod-id-direccion', value=direccion, placeholder='Dirección (máx. 200 caracteres)', debounce=True),
                html.P(nota_pie(direccion), style={'color': 'red', 'margin': '0px'}),
            ], style={'width': '50%'}),
            dbc.Col([
                html.H5(['Comuna']),
                dcc.Dropdown(op_comunas, id='mod-id-comuna', value=comuna),
                html.P(nota_pie(comunas[comuna]), style={'color': 'red', 'margin': '0px'}),
            ], style={'width': '18%'}),
        ]),
        linea,
    ])

def mod_fecha_visita(dia, datos):
    return html.Div([
        dbc.Row([
            dbc.Col([
                html.H5('Seleccione una fecha:'),
                dcc.DatePickerSingle(
                    id='mod-pick-fecha',
                    min_date_allowed=dia_laboral(),
                    max_date_allowed=fecha_final,
                    disabled_days=feriados,
                    first_day_of_week=1,
                    initial_visible_month=str(mes_sel),
                    date=dia,
                    display_format='D MMM YYYY',
                    stay_open_on_select=False,
                    show_outside_days=False,
                ),
                html.P(nota_pie(fecha_str(dia).strftime('%-d %b %Y')), style={'color': 'red', 'margin': '0px'}),
            ],
            style={'width': '18%'}
            ),
            dbc.Col([
                html.H5('Visitas programadas para dicha fecha'),
                dag.AgGrid(
                    id='mod-grid-agrega',
                    rowData=datos, # ***
                    defaultColDef={'resizable': True},
                    columnDefs=columnDefs_ing,
                    columnSize='sizeToFit',
                    dashGridOptions = {
                        'domLayout': 'autoHeight',
                        'rowSelection': 'single',
                    },
                    style={'width': '100%'},
                ),
            ],
            style={'width': '82%'}
            ),
        ]),
        linea,
    ])

def ajuste_fto_hora(hora):
    if hora == None:
        return None
    else:
        return horas_15[hora]

def mod_horario(inicio, termino, instal):
    return html.Div([
        dbc.Row([
            dbc.Col(
                html.H5(['Especifique el horario:']),
                style={'width': '18%'}
            ),
            dbc.Col([
                html.H5(['De inicio']),
                dcc.Dropdown(op_horas, id='mod-hr-inicio', value=inicio, style={'margin-right': '45px'}),
                html.P(nota_pie(ajuste_fto_hora(inicio)), style={'color': 'red', 'margin': '0px'}),
            ], style={'width': '15%'}),
            dbc.Col([
                html.H5(['De término']),
                dcc.Dropdown(op_horas, id='mod-hr-termino', value=termino, style={'margin-right': '45px'}),
                html.P(nota_pie(ajuste_fto_hora(termino)), style={'color': 'red', 'margin': '0px'}),
            ], style={'width': '15%'}),
            dbc.Col([
                html.H5(['De instalación']),
                dcc.Dropdown(op_horas, id='mod-hr-instala', value=instal, style={'margin-right': '45px'}),
                html.P(nota_pie(ajuste_fto_hora(instal)), style={'color': 'red', 'margin': '0px'}),
            ], style={'width': '15%'}),
        ]),
        linea,
    ])

def mod_contacto(ctto, tel_ctto, mail_ctto, cargo_ctto, ori, tel_ori, mail_ori):
    return html.Div([
        dbc.Row([
            dbc.Col(
                html.H5(['Añada información', html.Br(), 'de contacto:']),
                style={'width': '18%'}
            ),
            dbc.Col([
                html.H5(['Contacto']),
                dbc.Input(value=ctto, type='text', id='mod-contacto-nom', placeholder='Nombre (máx. 60 caracteres)', debounce=True),
                html.P(nota_pie(ctto), style={'color': 'red', 'margin': '0px'}),
                dbc.Input(value=tel_ctto, type='text', id='mod-contacto-cel', placeholder='Celular (máx. 20 caracteres)', debounce=True),
                html.P(nota_pie(tel_ctto), style={'color': 'red', 'margin': '0px'}),
                dbc.Input(value=mail_ctto, type='text', id='mod-contacto-mail', placeholder='Correo (máx. 50 caracteres)', debounce=True),
                html.P(nota_pie(mail_ctto), style={'color': 'red', 'margin': '0px'}),
                dbc.Input(value=cargo_ctto, type='text', id='mod-contacto-cargo', placeholder='Cargo (máx. 40 caracteres)', debounce=True),
                html.P(nota_pie(cargo_ctto), style={'color': 'red', 'margin': '0px'}),
            ], style={'width': '40%'}),
            dbc.Col([
                html.H5(['Orientador']),
                dbc.Input(value=ori, type='text', id='mod-orienta-nom', placeholder='Nombre (máx. 60 caracteres)', debounce=True),
                html.P(nota_pie(ori), style={'color': 'red', 'margin': '0px'}),
                dbc.Input(value=tel_ori, type='text', id='mod-orienta-cel', placeholder='Celular (máx. 20 caracteres)', debounce=True),
                html.P(nota_pie(tel_ori), style={'color': 'red', 'margin': '0px'}),
                dbc.Input(value=mail_ori, type='text', id='mod-orienta-mail', placeholder='Correo (máx. 50 caracteres)', debounce=True),
                html.P(nota_pie(mail_ori), style={'color': 'red', 'margin': '0px'}),
            ], style={'width': '40%'}),
        ]),
        linea,
    ])

# asistencia
def mod_asistencia(cantidad, detalle):
    return html.Div([
        dbc.Row([
            dbc.Col(
                html.H5('Asistencia estimada:'),
                style={'width': '18%'}
            ),
            dbc.Col([
                html.H5('Cantidad'),
                dbc.Input(type='text', id='mod-cantidad', value=cantidad, placeholder='Cantidad de alumnos/as (máx. 80 caracteres)', debounce=True),
                html.P(nota_pie(cantidad), style={'color': 'red', 'margin': '0px'}),
            ], style={'width': '30%'}),
            dbc.Col([
                html.H5('Detalle asistencia/programa'),
                dbc.Textarea(placeholder='Detalle de asistencia y programa', id='mod-cant-detalle', value=detalle,
                             style={'width': '97%', 'height': '4.5rem', 'border': '1px solid #d5d5d5', 'borderRadius': '5px'}),
                html.P(nota_pie(detalle), style={'color': 'red', 'margin': '0px'}),
            ], style={'width': '52%'}),
        ]),
        linea,
    ])

def mod_observaciones(obs):
    return html.Div([
        dbc.Row([
            dbc.Col([
                html.H5(['Observaciones:']),
                dbc.Textarea(id='mod-txt-observa', placeholder='Observaciones pertinentes de distinta índole', value=obs,
                             style={'width': '100%', 'height': '6rem', 'border': '1px solid #d5d5d5', 'borderRadius': '5px'}),
                html.P(nota_pie(obs), style={'color': 'red', 'margin': '0px'}),
            ], style={'width': '52%'}),
        ]),
        linea,
    ])

# botones
def mod_botones_modifica(fecha):
    return html.Div(
        dbc.Row([
            dbc.Col([
#                dbc.Button('Aplicar cambios', id='btn-mod-aplica', outline=True, color="primary", className='me-2', n_clicks=0, disabled=disabled_btn(fecha),
                dbc.Button('Aplicar cambios', id='btn-mod-aplica', outline=True, color="primary", className='me-2', n_clicks=0, style={'width': '15%'}),
                dbc.Button('Volver', id='btn-mod-volver', outline=True, color="primary", className='me-2', n_clicks=0, style={'width': '15%'}),
            ])
        ], style={'margin-bottom': '2rem'})
    )

# modal que informa que fecha no está disponible: modifica
mod_fecha_no_disponible = html.Div(
    dbc.Modal(
        [
            dbc.ModalHeader(html.H4('No es posible agregar visita'), close_button=False),
            dbc.ModalBody(html.Div('La fecha escogida ya no está disponible para agregar una nueva visita. Algún otro usuario la ocupó en el intertanto.')),
            dbc.ModalFooter(dbc.Button('Cerrar', id='cierra-fecha-no-disponible2', outline=True, color="primary", className='me-2', n_clicks=0, style={'width': '20%'})),
        ],
        id='modal-fecha-no-disponible2',
        size='lg',
        centered=True,
    ),
)

def form_modifica_detalle(ori, datos):
    return dbc.Form([
        html.H3(['Modificación de datos de visita'], className='titulo_pagina'),
        linea,
        identificacion(ori.nombre, ori.rbd),
        mod_estatus(ori.estatus),
        mod_direccion(ori.direccion, ori.comuna_id),
        mod_fecha_visita(ori.fecha, datos),
        mod_horario(ori.hora_ini, ori.hora_fin, ori.hora_ins),
        mod_contacto(ori.contacto, ori.contacto_tel, ori.contacto_mail, ori.contacto_cargo, ori.orientador, ori.orientador_tel, ori.orientador_mail),
        mod_asistencia(ori.asistentes, ori.asistentes_prog),
        mod_observaciones(ori.observaciones),
        mod_botones_modifica(ori.fecha),
        mod_fecha_no_disponible, # modal fecha no disponible
    ])

#### Forma
def form_modifica(datos):
    return html.Div([
        html.H3('Modifica/Elimina Visita', className='titulo_pagina'),
        linea,
        html.H5('En esta sección puede modificar y eliminar visitas no realizadas aún y confirmar si se realizaron visitas de fechas anteriores.',
               style={'fontWeight': 'normal', 'marginLeft': '10px'}),
        dbc.Row([
            dbc.Col([
                grid_modifica(datos),
                iconos_modifica,
                nota_modifica,
            ], style={'width': '75%'}),
            dbc.Col(
                botones_modifica, style={'width': '20%', 'marginBottom': '180px'}
            ),
        ], align='center'),
        confirma_realizacion,
        confirma_eliminacion,
    ])

### Invitaciones
opciones_asiste3 = {1: 'Aceptadas', 0: 'Rechazadas', 2: 'No confirmadas'}

def botones_condicion_asiste():
    return html.Div([
        dbc.Row([
            html.H5('Filtro según asistencia:', style={'width': '18%', 'display': 'flex'}),
            dcc.RadioItems(
                id = 'selec-op-asiste',
                options = opciones(opciones_asiste3),
                value = 2,
                inline = True,
                style = {'textAlign': 'left', 'width': '50%', 'display': 'flex', 'marginTop': '0px', 'padding': '0px'},
                labelStyle = {'display': 'inline-block', 'fontSize': '16px', 'fontWeight': 'normal'},
                inputStyle = {'marginRight': '5px', 'marginLeft': '20px'},
            )
        ]),
    ], style={'margin': '0px 0px 0px 10px'})

getRowStyle_invita = {
    "styleConditions": [
        {
            "condition": "params.data.invitacion == 0",
            "style": {"backgroundColor": "#ccccff"},
        },
    ],
    "defaultStyle": {"backgroundColor": "#fff"},
}

columnDefs_invitaciones = [
    {'field': 'orden', 'headerName': 'N', 'width': 55, 'cellStyle': {'textAlign': 'center'}},  ### crear una nueva variable con el orden local
    {'field': 'fecha', 'width': 150, 'cellStyle': {'textAlign': 'center'}},
    {'field': 'rbd', 'headerName': 'RBD', 'width': 100, 'cellStyle': {'textAlign': 'center'}, 'filter': True},
    {'field': 'nombre', 'headerName': 'Colegio', 'width': 420, 'filter': True},
    {'field': 'direccion', 'headerName': 'Dirección', 'width': 420, 'filter': True},
    {'field': 'comuna', 'filter': True},
]

def grid_invitaciones(datos=[]):
    return dag.AgGrid(
        id='grid-invita',
        rowData=datos,
        defaultColDef={'resizable': True},
        columnDefs=columnDefs_invitaciones,
        columnSize='sizeToFit',
        dashGridOptions = {
            'domLayout': 'autoHeight',
            'rowSelection': 'single',
        },
        getRowStyle=getRowStyle_invita,
        style={'width': '100%', 'margin': 0},
    )

def selector_asiste2(usuario=0, valor=2):
    return html.Div([
        html.H6('Asistencia a visita:', style={'fontSize': '17px', 'margin': 0}),
        dbc.Row([
            dbc.Col([
                html.P(f'{usuarios[usuario]}:', style={'fontSize': '16px', 'margin': '2px 0px 0px 20px'}),
            ], width='auto', style={'padding': 0}),
            dbc.Col([
                dcc.RadioItems(
                    id = 'selector-confirma-asist',
                    options=opciones_asiste,
                    value = valor,
                    inline = True,
                    labelStyle = {'display': 'inline-block', 'fontSize': '16px', 'fontWeight': 'normal'},
                    inputStyle = {'marginRight': '5px', 'marginLeft': '20px'},
                )
            ], width='auto', style={'padding': 0, 'marginTop': '2px'}),
            dbc.Col(
                dbc.Button('Aplicar cambio', id='btn-confirma-asist', outline=True, color="primary", className='me-2', n_clicks=0, disabled=True,
                           style={'fontSize': '14px', 'padding': '4px 28px', 'marginTop': '6px', 'marginLeft': '8px'})
            )
        ]),
    ], style={'marginLeft': 10, 'marginBottom': 1, 'marginTop': 20})

# contenido del modal
def contenido_modal_invita(datos_visita={}, datos_dic={}, usuario=0, asiste=2):
    return html.Div([
        dbc.Row([
            dbc.Col([
                html.Div(selector_asiste2(usuario, asiste), id='opciones-modal-invita'),
                linea,
                dbc.Accordion(
                    [
                        dbc.AccordionItem(
                            [
                                seccion_info_gral(datos_visita),
                            ],
                            title="1. Antecedentes de la visita",
                        ),
                        dbc.AccordionItem(
                            [
                                html.Div(seccion_universidades_asisten2(datos_dic), id='sec-univ-asist'),
                            ],
                            title="2. Asistencia de universidades",
                        ),
                    ],
                )
            ])
        ])
    ])

# modal que confirma asisitencia
confirma_asistencia = html.Div(
    dbc.Modal(
        [
            dbc.ModalHeader(html.H4('Confirma asistencia a visita', style={'fontSize': '2rem', 'marginBottom': '0px'}), close_button=False),
            dbc.ModalBody(contenido_modal_invita(), id='confirma-asist-contenido', style={'paddingTop': '0px'}),
            dbc.ModalFooter(
                html.Div([
                    dbc.Button('Cerrar', id='btn-cerrar-confirma-asist', outline=True, color="primary", className='me-2', n_clicks=0, style={'padding': '6px 25px'}),
                ])
            ),
        ],
        id='modal-confirma-asist',
        size='lg',
        keyboard=False,
        backdrop="static",
   ),
)

icono_invita = html.Div(
    [
        html.Img(src='./assets/morado3.png', height='25px'),
        html.P('Organizadas por el usuario', style=estilo_iconos),
    ],
    id='icono-invita',
    hidden=False,
    style={'margin': '10px 0px'}
)

#### Forma
def form_invitaciones(datos=[]):
    return html.Div([
        html.H3('Invitaciones Recibidas', className='titulo_pagina'),
        html.H6('(Se incluyen las organizadas por el propio usuario)', style={'margin': '-10px 10px'}),
        linea,
        dbc.Row([
            dbc.Col([
                botones_condicion_asiste(),
                grid_invitaciones(datos),
                icono_invita,
            ], style={'width': '100%', 'margin': 0}
            ),
        ], style={'marginBottom': '30px'}),
        confirma_asistencia,
    ])


### Footer
footer_container = dbc.Container(
    dbc.Row(
        dbc.Col(
            html.Footer(
                html.P(['2026:  Corporación de Universidades Privadas'], style={'margin': 0}),
                style={'background': '#175a96', 'color': 'white', 'padding': '15px 25px', 'font-size': '20px'}
            )
        )
    ),
    className="mt-auto"
)

### Contenido
def contenido():
    return html.Div(
        [
            html.Div([], id='advert', hidden=True),
            html.Div(form_inicio(fn_programadas_visita(programadas.tra_dic, mes=fn_mes()), fn_mes()), id='resumen', hidden=False),
            html.Div(form_agrega(dia_laboral(), fn_programadas_fecha(programadas.tra_dic, dia_laboral())), id='agrega', hidden=True),
            html.Div([], id='elimina', hidden=True),
            html.Div(form_invitaciones(), id='invita', hidden=True),
            html.Div([], id='detalle', hidden=True),
        ],
        id='contenido'
    )

### Layout
parametros = {
    'usuario': usuario,
    'origen': 'btn-resumen',
    'destino': '',
    'origen_vista': 'resumen',
    'mes': max(dia_laboral(), fecha_inicial).month,
    'fecha_seleccionada': max(dia_laboral(), fecha_inicial).strftime('%Y-%m-%d'),
    'id_visita': None,
    'id_visita_mod': None,
    'fecha_mod_seleccionada': None,
    'estatus': None,
    'abre_detalle': False,
    'dic_modifica': {},  # debiera desaparece
    'datos_previos': {},
    'sel_asiste': None,
    'cond_asiste': 2,
}

estilo_layout = {
    "display": "flex",
    "flexDirection": "column",
    "minHeight": "100vh"  # Full viewport height
}

def serve_layout():
    return html.Div([
        dbc.Container([
            navbar,
            modal_acceso,
            html.Div([], id='advert', hidden=True),
            html.Div(form_inicio(fn_programadas_visita(programadas.tra_dic, mes=fn_mes()), fn_mes()), id='resumen', hidden=False),
            html.Div(form_agrega(dia_laboral(), fn_programadas_fecha(programadas.tra_dic, dia_laboral())), id='agrega', hidden=True),
            html.Div([], id='elimina', hidden=True),
            html.Div(form_invitaciones(), id='invita', hidden=True),
            html.Div([], id='detalle', hidden=True),
        ],
        className="flex-grow-1"  # Takes up available space
        ),
        footer_container,

        dcc.Store(id='datos', data=programadas.tra_dic),
        dcc.Store(id='asisten', data=asisten.ori_dic),
        dcc.Store(id='bloqueadas', data=obtiene_bloqueadas()),
        dcc.Store(id='parametros', data=parametros),
    ],
    style=estilo_layout
)

### Adiciones
lista_btn = [
    'btn-resumen',
    'btn-agrega',
    'btn-elimina',
    'btn-invita',
]

dic_color = dict(zip(lista_btn, [dash.no_update]*len(lista_btn)))

lista_vistas = [
    'resumen',
    'agrega',
    'elimina',
    'detalle',
    'invita',
]

lista_vistas_red = lista_vistas.copy()
lista_vistas_red.remove('detalle')

indice = lambda x: lista_btn.index(x)
base_retorno = lambda n: [dash.no_update]*n

def opcion_visible(disparador, param, lista):
    dic = {param['origen_vista']: True}  # inicializa un diccionario
    dic_color_local = dic_color.copy()
    if disparador != param['origen']:
        dic_color_local.update({disparador: 'componente-inicio', param['origen']: 'componentes'})
    origen_v = ''

    match disparador:
        case 'btn-resumen':
            dic['resumen'] = False
            origen_v = 'resumen'
        case 'btn-agrega':
            dic['agrega'] = False
            origen_v = 'agrega'
        case 'btn-elimina':
            if param['abre_detalle']:
                dic['detalle'] = False
                origen_v = 'detalle'
            else:
                dic['elimina'] = False
                origen_v = 'elimina'
        case 'btn-invita':
            dic['invita'] = False
            origen_v = 'invita'

    lista_vista = [dic.get(x, dash.no_update) for x in lista]
    lista_color = [dic_color_local.get(x, dash.no_update) for x in lista_btn]
    param.update({'origen_vista': origen_v})

    return lista_vista, lista_color, param

### Ejecución
estilo = {
    'href':'./assets/style_2026.css',
    'rel':'stylesheet',
}

app = Dash(__name__, prevent_initial_callbacks=True, external_stylesheets=[dbc.themes.CERULEAN, estilo])

app.config.suppress_callback_exceptions = True

app.layout = serve_layout
app.title = 'Visitas CUP 2026'

server = app.server

# CALLBACK
# 1 --------------------------------------------------------------------------
# inicio: selección de opción (usuario acreditado: cambio de visualización)
@app.callback(
    Output('parametros', 'data', allow_duplicate=True),
    Output('modal-acceso', 'is_open', allow_duplicate=True),

    Output('resumen', 'hidden', allow_duplicate=True),
    Output('agrega', 'hidden', allow_duplicate=True),
    Output('elimina', 'hidden', allow_duplicate=True),
    Output('detalle', 'hidden', allow_duplicate=True),
    Output('invita', 'hidden', allow_duplicate=True),

    Output('btn-resumen', 'className', allow_duplicate=True),
    Output('btn-agrega', 'className', allow_duplicate=True),
    Output('btn-elimina', 'className', allow_duplicate=True),
    Output('btn-invita', 'className', allow_duplicate=True),

    Input('btn-resumen', 'n_clicks'),
    Input('btn-agrega', 'n_clicks'),
    Input('btn-elimina', 'n_clicks'),
    Input('btn-invita', 'n_clicks'),
    State('parametros', 'data'),
)
def selecciona_opcion(click_1, click_2, click_3, click_4, param):
    disparador = dash.ctx.triggered_id

    if bool(param['usuario']): # ha ingresado
        if param['origen'] == disparador:
            return dash.no_update, dash.no_update, *base_retorno(9)
        else:
            lista_vista, lista_color, parametros = opcion_visible(disparador, param, lista_vistas)
            parametros.update({'origen': disparador, 'destino': ''})
            return parametros, dash.no_update, *lista_vista, *lista_color
    else:
        param.update({'destino': disparador})
        return param, True, *base_retorno(9)
        
# 2 --------------------------------------------------------------------------
# acceso: ingreso de usuario (actualización de páginas: funciones que actualicen)
@app.callback(
    Output('parametros', 'data', allow_duplicate=True),
    Output('loc-usuario', 'children'),
    Output('modal-acceso', 'is_open', allow_duplicate=True),

    Output('advert', 'children'),
    Output('elimina', 'children'),
    Output('invita', 'children'),

    Output('advert', 'hidden', allow_duplicate=True),
    Output('resumen', 'hidden', allow_duplicate=True),
    Output('agrega', 'hidden', allow_duplicate=True),
    Output('elimina', 'hidden', allow_duplicate=True),
    Output('invita', 'hidden', allow_duplicate=True),

    Output('btn-resumen', 'className', allow_duplicate=True),
    Output('btn-agrega', 'className', allow_duplicate=True),
    Output('btn-elimina', 'className', allow_duplicate=True),
    Output('btn-invita', 'className', allow_duplicate=True),

    Input('btn-ingresar', 'n_clicks'),
    State('sel-u', 'value'),
    State('inp-pw', 'value'),
    State('datos', 'data'),
    State('asisten', 'data'),
    State('parametros', 'data'),
)
def acreditacion_usuario(click, universidad, contrasena, datos, datos_asisten, param):
    if contrasena is None:
        return dash.no_update, dash.no_update, False, *base_retorno(12)
    elif os.environ.get('U'+str(universidad)) == contrasena:
        data = DatosInicio(datos, datos_asisten, universidad)

        data_modifica, cant_realizada = data.programadas_usuario()
        data_invita = data.invitaciones(param['cond_asiste'])
        cant_invita = data.cantidad_visitas()

        texto = data.texto_advertencia()
        texto_hd = not bool(cant_invita or cant_realizada)

        contenido = [form_advertencia(texto), form_modifica(data_modifica), form_invitaciones(data_invita)]
        nombre_usuario = universidades[universidad]
        lista_vista, lista_color, param = opcion_visible(param['destino'], param, lista_vistas_red)

        param.update({'usuario': universidad, 'origen': param['destino'], 'destino': ''})

        return param, nombre_usuario, False, *contenido, texto_hd, *lista_vista, *lista_color
    else:
        return dash.no_update, dash.no_update, False, *base_retorno(12)
    

# acceso: cancela ingreso
@app.callback(
    Output('modal-acceso', 'is_open', allow_duplicate=True),
    Input('btn-cancelar', 'n_clicks'),
)
def cancela_acceso(click):
    return False

# 3 --------------------------------------------------------------------------
# cambio de mes:
@app.callback(
    Output('grid-programadas', 'rowData', allow_duplicate=True),
    Output('parametros', 'data', allow_duplicate=True),
    Input('selec-mes', 'value'),
    State('datos', 'data'),
    State('parametros', 'data'),
)
def programadas_visita_mes(mes, datos, param):
    param['mes'] = mes
    return fn_programadas_visita(datos, mes=mes), param

# cambio de día: 
@app.callback(
    Output('grid-agrega', 'rowData', allow_duplicate=True),
    Output('btn-agregar-visita', 'disabled', allow_duplicate=True),
    Output('parametros', 'data', allow_duplicate=True),
    Input('pick-fecha', 'date'),
    State('datos', 'data'),
    State('bloqueadas', 'data'),
    State('parametros', 'data'),
)
def ferias_programadas_fecha(fecha, datos, bloq, param):
    param['fecha_seleccionada'] = fecha
    return fn_programadas_fecha(datos, fecha), en_bloqueadas2(fecha, bloq), param  # no necesita actualizar las bloqueadas

# --------------------------------------------------------------------------
# abre modal reporte
@app.callback(
    Output('reporte-prog-contenido', 'children'),
    Output('modal-reporte-prog', 'is_open', allow_duplicate=True),
    Output('grid-programadas', 'selectedRows'),
#    Output('descarga-reporte', 'disabled'),
    Output('div-descarga-reporte', 'hidden'),
    Output('parametros', 'data', allow_duplicate=True),
    Input('grid-programadas', 'selectedRows'),
    State('datos', 'data'),
    State('asisten', 'data'),
    State('parametros', 'data'),
)
def abre_modal_reporte(fila, datos, datos_asisten, param):
    if (fila == None) | (fila == []):
        raise PreventUpdate
    else:
        id_sel = fila[0]['programada_id']
        datos_dic = next(item for item in datos if item['programada_id'] == id_sel)
        asisten = fn_universidades_asisten(datos_asisten, id_sel)
        uasiste = fn_usuario_asiste(datos_asisten, param['usuario'], id_sel)
        fecha = datos_dic['fecha']
        param.update({'sel_asiste': uasiste, 'id_visita': id_sel})
        
        return escoge_reporte(param['usuario'], datos_dic, asisten, uasiste, fecha), True, [], (not bool(param['usuario'])), param

# cierra modal reporte
@app.callback(
    Output('modal-reporte-prog', 'is_open', allow_duplicate=True),
    Input('btn-cerrar-reporte-prog', 'n_clicks'),
)
def cierra_modal_reporte(click):
    return False

# descarga reporte de la visita en formato pdf
@app.callback(
    Output('descarga-reporte-archivo', 'data'),
    Input('descarga-reporte', 'n_clicks'),
    State('datos', 'data'),
    State('asisten', 'data'),
    State('parametros', 'data'),
)
def descarga_reporte_pdf(click, datos, datos_asisten, param):
    id_sel = param['id_visita']
    datos_dic = next(item for item in datos if item['programada_id'] == id_sel)
    asisten = fn_universidades_asisten(datos_asisten, id_sel).get(1, None)
    doc = exporta_reporte(datos_dic, asisten)

    return dcc.send_bytes(doc, f"reporte_{str(datos_dic['rbd'])}.pdf")

# --------------------------------------------------------------------------
# abre modal calendario
@app.callback(
    Output(component_id='calendario-contenido', component_property='src', allow_duplicate=True),
    Output('modal-calendario', 'is_open', allow_duplicate=True),
    Output('selec-mes-cal', 'value'),
    Input('abre-calendario', 'n_clicks'),
    State('datos', 'data'),
    State('parametros', 'data'),
)
def abre_modal_calendario(click, datos, param):
    contenido = crea_figura(fn_programadas(datos).collect(), max(3, param['mes']))
    return contenido, True, max(3, param['mes'])

# cierra modal calendario
@app.callback(
    Output('modal-calendario', 'is_open', allow_duplicate=True),
    Input('btn-cerrar-calendario', 'n_clicks'),
)
def cierra_modal_calendario(click):
    return False

# cambia de mes en modal calendario
@app.callback(
    Output(component_id='calendario-contenido', component_property='src', allow_duplicate=True),
    Input('selec-mes-cal', 'value'),
    State('datos', 'data'),
)
def cambia_mes_calendario(mes, datos):
    contenido = crea_figura(fn_programadas(datos).collect(), mes)
    return contenido

# imprimir calendario
@app.callback(
    Output('exporta-calendario', 'data'),
    Input('imprimir-calendario', 'n_clicks'),
    State('selec-mes-cal', 'value'),
    State('datos', 'data'),
)
def imprimir_calendario(click, mes, datos):
    contenido = crea_figura(fn_programadas(datos).collect(), mes).removeprefix("data:image/png;base64,")
    contenido_bytes = base64.b64decode(contenido)
    output = io.BytesIO(contenido_bytes)
    output.seek(0)

    return dcc.send_bytes(output.read(), f"calendario_{meses[mes-1].lower()}_2026.png")

# --------------------------------------------------------------------------
# exporta visitas programadas a excel
@app.callback(
    Output('exporta-visitas-archivo', 'data'),
    Input('exporta-visitas', 'n_clicks'),
    State('datos', 'data'),
    State('parametros', 'data'),
)
def exporta_visitas_excel(click, datos, param):
    df = fn_exporta_programada(datos, param['usuario'], param['mes'])
    if param['usuario'] == 0:
        return dcc.send_bytes(df, 'visitas.xlsx')
    else:
        return dcc.send_bytes(df, 'visitas_detalle.xlsx')

# --------------------------------------------------------------------------
# selección de RBD y nombre
@app.callback(
    Output('sel-rbd', 'value'),
    Output('sel-nombre', 'value'),
    Output('id-comuna', 'value'),
    Input('sel-rbd', 'value'),
    Input('sel-nombre', 'value'),
    Input('btn-limpiar-sel', 'n_clicks'),
)
def completa_rbd_y_nombre(rbd, nombre, click):

    disparador = dash.ctx.triggered_id

    match disparador:
        case 'sel-rbd':
            try:
                return dash.no_update, rbd, colegios_comuna[rbd]
            except:
                return None, dash.no_update, dash.no_update
        case 'sel-nombre':
            return nombre, dash.no_update, colegios_comuna[nombre]
        case 'btn-limpiar-sel':
            return None, None, None

# limpia todos los campos
@app.callback(
    Output('agrega', 'children', allow_duplicate=True),
    Input('btn-limpiar', 'n_clicks'),
    State('datos', 'data'),
    State('parametros', 'data'),
)
def limpia_todos_los_campos(click, datos, param):
    fecha = a_fecha(param['fecha_seleccionada'])
    return form_agrega(fecha, fn_programadas_fecha(datos, fecha))

# --------------------------------------------------------------------------
# ingreso de visita
@app.callback(
    Output('datos', 'data', allow_duplicate=True),
    Output('asisten', 'data', allow_duplicate=True),
    Output('bloqueadas', 'data', allow_duplicate=True),

    Output('agrega', 'children', allow_duplicate=True),
    
    Output('modal-fecha-no-disponible', 'is_open'),
    Output('btn-agregar-visita', 'disabled', allow_duplicate=True),

    Output('grid-programadas', 'rowData', allow_duplicate=True),    # programadas
    Output('grid-agrega', 'rowData', allow_duplicate=True),    # agrega
    Output('grid-modifica', 'rowData', allow_duplicate=True),   # modifica
    Output('grid-invita', 'rowData', allow_duplicate=True),     # invitaciones
    
    Input('btn-agregar-visita', 'n_clicks'),
    
    State('pick-fecha', 'date'),
    State('sel-rbd', 'value'),
    State('sel-nombre', 'value'),
    State('id-direccion', 'value'),
    State('id-comuna', 'value'),
    State('hr-inicio', 'value'),
    State('hr-termino', 'value'),
    State('hr-instala', 'value'),
    State('contacto-nom', 'value'),
    State('contacto-cel', 'value'),
    State('contacto-mail', 'value'),
    State('contacto-cargo', 'value'),
    State('orienta-nom', 'value'),
    State('orienta-cel', 'value'),
    State('orienta-mail', 'value'),
    State('cantidad', 'value'),
    State('cant-detalle', 'value'),
    State('def-estatus', 'value'),
    State('txt-observa', 'value'),
    State('parametros', 'data'),
)
def agrega_visita(click, fecha, rbd, nombre, direccion, comuna, inicio, termino, instala, contacto, contac_cel, contac_mail, contac_cargo, orienta, orienta_cel,
                  orienta_mail, cantidad, cant_detalle, estatus, observa, param):

    if not rbd:
        raise PreventUpdate
    else:
        bloqueado = en_bloqueadas(fecha)  # aquí sí se necesita verificar si la fecha está bloqueada
    
        if bloqueado:
            data = Actualiza(param)
            return data.programadas_dic(), data.asisten_dic(), data.bloqueadas, dash.no_update, True, True, data.programadas_visita(), data.programadas_fecha(), dash.no_update, data.invitaciones(param['cond_asiste'])  # cambio aquí
        else:
            us = param['usuario']
            lista = [us, universidades[us], a_fecha(fecha), rbd, colegios[nombre], direccion, comuna, a_hora(inicio), a_hora(termino), a_hora(instala), contacto, contac_cel,
                     contac_mail, contac_cargo, orienta, orienta_cel, orienta_mail, cantidad, cant_detalle, estatus, observa]
        
            tupla_ = Nt_programada(*lista)
            dic = {k: v for k, v in tupla_._asdict().items() if v}
            tupla = Nt_programada(**dic)
    
            nueva_programada(tupla)
            data = Actualiza(param)
    
            return data.programadas_dic(), data.asisten_dic(), data.bloqueadas, form_agrega(a_fecha(fecha), data.programadas_fecha()), False, False, data.programadas_visita(), data.programadas_fecha(), data.programadas_usuario()[0], dash.no_update


# cierra modal: fecha no disponible (agrega)
@app.callback(
    Output('modal-fecha-no-disponible', 'is_open', allow_duplicate=True),
    Input('cierra-fecha-no-disponible', 'n_clicks'),
)
def cierra_modal_fecha_no_disponible(click):
    return False

# --------------------------------------------------------------------------
# selecciona visita para modificar o eliminar / abre modal que confirma realización de visita
@app.callback(
    Output('modal-confirma-visita', 'is_open', allow_duplicate=True),
    Output('mod-selector-asiste', 'value'),
    Output('parametros', 'data', allow_duplicate=True),
    Output('btn-mod-visita', 'disabled'),
    Output('btn-elim-visita', 'disabled'),
    Input('grid-modifica', 'selectedRows'),
    State('datos', 'data'),
    State('parametros', 'data'),
)
def modal_confirma_realizacion(fila, datos, param):
    if fila == None or fila == []:
        raise PreventUpdate
    else:
        visita_sel = fila[0]['programada_id']
        datos_dic = next(item for item in datos if item['programada_id'] == visita_sel)
        param.update({'id_visita_mod': visita_sel, 'estatus': datos_dic['estatus'], 'fecha_mod_seleccionada': datos_dic['fecha']})
        tipo = fila[0]['tipo']
        
        if tipo in [0, 3]:
            return False, dash.no_update, param, False, False
        elif tipo in [1, 2]:
            return True, datos_dic['estatus'], param, True, True

# cierra modal que confirma realización de visita: actualiza
@app.callback(
    Output('datos', 'data', allow_duplicate=True),
    Output('asisten', 'data', allow_duplicate=True),
    Output('bloqueadas', 'data', allow_duplicate=True),

    Output('modal-confirma-visita', 'is_open', allow_duplicate=True),
    Output('grid-modifica', 'selectedRows'),

    Output('txt-advert', 'children', allow_duplicate=True),
    Output('advert', 'hidden', allow_duplicate=True),
    Output('grid-programadas', 'rowData', allow_duplicate=True),
    Output('grid-modifica', 'rowData', allow_duplicate=True),
    
    Input('cambia-confirma-visita', 'n_clicks'),
    Input('btn-cerrar-confirma-visita', 'n_clicks'),

    State('mod-selector-asiste', 'value'),
    State('datos', 'data'),
    State('parametros', 'data'),
    State('grid-modifica', 'selectedRows'),
)
def cierra_modal_confirma_realizacion(click1, click2, nuevo_estatus, datos, param, lista):
    disparador = dash.ctx.triggered_id

    match disparador:
        case 'cambia-confirma-visita':
            if nuevo_estatus != param['estatus']:
                cambia_estatus(param['id_visita_mod'], nuevo_estatus)
                data = Actualiza(param)
                return data.programadas_dic(), data.asisten_dic(), data.bloqueadas, False, [], data.texto_advertencia(), data.oculta_advertencia(), data.programadas_visita(), data.programadas_usuario()[0]
            else:
                return dash.no_update, dash.no_update, dash.no_update, False, [], *[dash.no_update]*4
        case 'btn-cerrar-confirma-visita':
            return dash.no_update, dash.no_update, dash.no_update, False, [], *[dash.no_update]*4

# --------------------------------------------------------------------------
# abre página para modificar visita
@app.callback(
    Output('detalle', 'children'),
    Output('detalle', 'hidden', allow_duplicate=True),
    Output('elimina', 'hidden', allow_duplicate=True),
    Output('parametros', 'data', allow_duplicate=True),
    Input('btn-mod-visita', 'n_clicks'),
    State('grid-modifica', 'selectedRows'),
    State('datos', 'data'),
    State('parametros', 'data'),
)
def abre_modifica_detalle(click, fila, datos, param):
    if (fila == []) | (fila == None):
        raise PreventUpdate
    else:
        visita_sel = fila[0]['programada_id']
        datos_dic = next(item for item in datos if item['programada_id'] == visita_sel)
        datos_dic_reducidos = dic_reducido(datos_dic)
        datos_fecha = fn_programadas_fecha(datos, datos_dic['fecha']) # ***
        param.update({'origen_vista': 'detalle', 'abre_detalle': True, 'dic_modifica': {}, 'datos_previos': datos_dic})
        return form_modifica_detalle(Nt_programada(**datos_dic_reducidos), datos_fecha), False, True, param

# cambia fecha en ventana de modifica
@app.callback(
    Output('mod-grid-agrega', 'rowData', allow_duplicate=True),
    Output('btn-mod-aplica', 'disabled', allow_duplicate=True),
    Input('mod-pick-fecha', 'date'),
    State('datos', 'data'),
    State('bloqueadas', 'data'),
    State('parametros', 'data'),
)
def mod_cambia_fecha(fecha, datos, bloq, param):
    return fn_programadas_fecha(datos, fecha), en_bloqueadas2(fecha, bloq, param['fecha_mod_seleccionada'])

# actualiza datos modificados
@app.callback(
    Output('detalle', 'hidden', allow_duplicate=True),
    Output('elimina', 'hidden', allow_duplicate=True),
    Output('parametros', 'data', allow_duplicate=True),
    
    Output('datos', 'data', allow_duplicate=True),
    Output('asisten', 'data', allow_duplicate=True),
    Output('bloqueadas', 'data', allow_duplicate=True),
    Output('txt-advert', 'children', allow_duplicate=True),
    Output('advert', 'hidden', allow_duplicate=True),
    Output('mod-grid-agrega', 'rowData', allow_duplicate=True),
    Output('grid-programadas', 'rowData', allow_duplicate=True),
    Output('grid-modifica', 'rowData', allow_duplicate=True),

    Output('modal-fecha-no-disponible2', 'is_open', allow_duplicate=True),
    Output('btn-mod-aplica', 'disabled', allow_duplicate=True),

    Input('btn-mod-aplica', 'n_clicks'),
    
    State('mod-pick-fecha', 'date', allow_optional=True),
    State('mod-id-direccion', 'value',  allow_optional=True),
    State('mod-id-comuna', 'value',  allow_optional=True),
    State('mod-hr-inicio', 'value',  allow_optional=True),
    State('mod-hr-termino', 'value',  allow_optional=True),
    State('mod-hr-instala', 'value',  allow_optional=True),
    State('mod-contacto-nom', 'value',  allow_optional=True),
    State('mod-contacto-cel', 'value',  allow_optional=True),
    State('mod-contacto-mail', 'value',  allow_optional=True),
    State('mod-contacto-cargo', 'value',  allow_optional=True),
    State('mod-orienta-nom', 'value',  allow_optional=True),
    State('mod-orienta-cel', 'value',  allow_optional=True),
    State('mod-orienta-mail', 'value',  allow_optional=True),
    State('mod-cantidad', 'value',  allow_optional=True),
    State('mod-cant-detalle', 'value',  allow_optional=True),
    State('mod-def-estatus', 'value',  allow_optional=True),
    State('mod-txt-observa', 'value',  allow_optional=True),
    State('parametros', 'data'),
)
def registra_modificacion(click, fecha, direccion, comuna, inicio, termino, instala, contacto, contac_cel, contac_mail, contac_cargo, orienta,
                          orienta_cel, orienta_mail, cantidad, cant_detalle, estatus, observa, param):

    lista_variables = [fecha, direccion, comuna, inicio, termino, instala, contacto, contac_cel, contac_mail, contac_cargo, orienta, orienta_cel,
                        orienta_mail, cantidad, cant_detalle, estatus, observa]
    lista_variables = list(map(lambda x: x if x != '' else None, lista_variables))

    dic_cambios = Nt_modifica(*lista_variables)._asdict()
    diferencia = dict(set(list(dic_cambios.items())).difference(list(param['datos_previos'].items())))

    def fn_retorna():
        dic_ajustado = {k: dic_fn_modifica.get(k, lambda x: x)(v) for k, v in diferencia.items()}
        modifica_programada(param['id_visita_mod'], dic_ajustado)

        data = Actualiza(param)
        param.update({'origen_vista': 'elimina', 'abre_detalle': False})
        return True, False, param, data.programadas_dic(), data.asisten_dic(), data.bloqueadas, data.texto_advertencia(), data.oculta_advertencia(), data.programadas_fecha(fecha), data.programadas_visita(), data.programadas_usuario()[0], False, dash.no_update
        
    if (diferencia == {}):  # | (click == 0):
        raise PreventUpdate
    else:
        fecha_loc = fecha_str(diferencia.get('fecha', None))
        if fecha_loc:
            if en_bloqueadas(fecha_loc, param['datos_previos']['fecha']):
                data = Actualiza(param)
                return dash.no_update, dash.no_update, dash.no_update, data.programadas_dic(), data.asisten_dic(), data.bloqueadas, *[dash.no_update]*2, data.programadas_fecha(fecha), data.programadas_visita(), dash.no_update, True, True
            else:
                return fn_retorna()
        else:
            return fn_retorna()

# cierra ventana de modifica sin actualizar datos
@app.callback(
    Output('detalle', 'hidden', allow_duplicate=True),
    Output('elimina', 'hidden', allow_duplicate=True),
    Output('parametros', 'data', allow_duplicate=True),
    Input('btn-mod-volver', 'n_clicks'),
    State('parametros', 'data'),
)
def cierra_sin_modificar(click, param):
    if click == 0:
        raise PreventUpdate
    else:
        param.update({'origen_vista': 'elimina', 'abre_detalle': False})
        return True, False, param

# cierra modal que informa que fecha no está disponible
@app.callback(
    Output('modal-fecha-no-disponible2', 'is_open', allow_duplicate=True),
    Input('cierra-fecha-no-disponible2', 'n_clicks'),
)
def cierra_modal_fecha_no_disponible2(click):
    return False

    # abre modal que solicita confirmación de eliminación de visita: 
@app.callback(
    Output('modal-confirma-elimina', 'is_open', allow_duplicate=True),
    Input('btn-elim-visita', 'n_clicks'),
)
def abre_modal_elimina(click):
    return True

# --------------------------------------------------------------------------
# cambia condición de asistencia en modal de reporte: 
@app.callback(
    Output('btn-selector-asiste', 'disabled', allow_duplicate=True),
    Input('selector-asiste', 'value'),
    State('parametros', 'data'),
)
def selecciona_opcion_asiste(seleccion, param):
    if seleccion == param['sel_asiste']:
        return True
    else:
        return False

# aplica cambio en asistencia: actualiza (datos_asiste, texto, cinta, reporte, invitaciones)
@app.callback(
    Output('asisten', 'data', allow_duplicate=True),

    Output('rep-asiste', 'children'),
    Output('rep-no-asiste', 'children'),
    Output('rep-no-confirma', 'children'),

    Output('txt-advert', 'children', allow_duplicate=True),
    Output('advert', 'hidden', allow_duplicate=True),
    Output('grid-invita', 'rowData', allow_duplicate=True),

    Output('btn-selector-asiste', 'disabled', allow_duplicate=True),
    
    Input('btn-selector-asiste', 'n_clicks'),
    State('selector-asiste', 'value'),
    State('parametros', 'data'),
)
def cambia_opcion_asiste(click, asiste, param):
    if click == 0:
        raise PreventUpdate
    else:
        modifica_asiste(param['usuario'], param['id_visita'], asiste)
    
        data = Actualiza(param)
        dic = data.universidades_asisten(param['id_visita'])
        ch_asiste = seccion_universidades_asisten(dic, 1)
        ch_no_asiste = seccion_universidades_asisten(dic, 0)
        ch_no_confirma = seccion_universidades_asisten(dic, 2)

        return data.asisten_dic(), ch_asiste, ch_no_asiste, ch_no_confirma, data.texto_advertencia(), data.oculta_advertencia(), data.invitaciones(param['cond_asiste']), True

# desiste eliminar visita
@app.callback(
    Output('modal-confirma-elimina', 'is_open', allow_duplicate=True),
    Input('btn-confirma-no-elimina', 'n_clicks'),
)
def confirma_no_elimina(click):
    return False

# elimina definitivamente visita: Actualiza
@app.callback(
    Output('modal-confirma-elimina', 'is_open', allow_duplicate=True),

    Output('datos', 'data', allow_duplicate=True),
    Output('asisten', 'data', allow_duplicate=True),
    Output('bloqueadas', 'data', allow_duplicate=True),
    Output('grid-programadas', 'rowData', allow_duplicate=True),
    Output('grid-agrega', 'rowData', allow_duplicate=True),
    Output('grid-modifica', 'rowData', allow_duplicate=True),

    Input('btn-confirma-elimina', 'n_clicks'),
    State('parametros', 'data'),
)
def confirma_elimina(click, param):
    if click == 0:
        raise PreventUpdate
    else:
        elimina_programada(param['id_visita_mod'])
        data = Actualiza(param)
        return False, data.programadas_dic(), data.asisten_dic(), data.bloqueadas, data.programadas_visita(), data.programadas_fecha(), data.programadas_usuario()[0]

# --------------------------------------------------------------------------
# navegación entre opciones de asistencia
@app.callback(
    Output('grid-invita', 'rowData', allow_duplicate=True),
    Output('parametros', 'data', allow_duplicate=True),
    Output('icono-invita', 'hidden'),
    Input('selec-op-asiste', 'value'),
    State('datos', 'data'),
    State('asisten', 'data'),
    State('parametros', 'data'),
)
def opciones_de_asistencia(opcion, datos, datos_asisten, param):
    ico = True if opcion == 2 else False
    param.update({'cond_asiste': opcion})
    return fn_invitaciones(datos, datos_asisten, param['usuario'], opcion), param, ico

# abre modal que confirma asistencia
@app.callback(
    Output('confirma-asist-contenido', 'children'),
    Output('modal-confirma-asist', 'is_open', allow_duplicate=True),
    Output('parametros', 'data', allow_duplicate=True),
    Input('grid-invita', 'selectedRows'),
    State('datos', 'data'),
    State('asisten', 'data'),
    State('parametros', 'data'),
)
def abre_modal_confirma_asist(fila, datos, datos_asisten, param):
    if (fila == None) | (fila == []):
        raise PreventUpdate
    else:
        id_sel = fila[0]['programada_id']
        datos_dic = next(item for item in datos if item['programada_id'] == id_sel)
        asisten = fn_universidades_asisten(datos_asisten, id_sel)
        uasiste = fn_usuario_asiste(datos_asisten, param['usuario'], id_sel)
        param.update({'sel_asiste': uasiste, 'id_visita': id_sel})
       
        return contenido_modal_invita(datos_dic, asisten, param['usuario'], uasiste), True, param

# cambia opción en modal de asistencia
@app.callback(
    Output('btn-confirma-asist', 'disabled'),
    Input('selector-confirma-asist', 'value'),
    State('parametros', 'data'),
)
def selecciona_opcion_asiste_ventana(seleccion, param):
    if seleccion == param['sel_asiste']:
        return True
    else:
        return False

# cierra modal que confirma asistencia
@app.callback(
    Output('modal-confirma-asist', 'is_open', allow_duplicate=True),
    Output('grid-invita', 'selectedRows'),
    Input('btn-cerrar-confirma-asist', 'n_clicks'),
)
def cierra_modal_confirma_asist(click):
    if click == 0:
        raise PreventUpdate
    else:
        return False, []

# selecciona opción de asistencia: Actualiza (datos_asiste, texto, cinta, invitaciones)
@app.callback(
    Output('asisten', 'data', allow_duplicate=True),
    Output('txt-advert', 'children', allow_duplicate=True),
    Output('advert', 'hidden', allow_duplicate=True),
    Output('grid-invita', 'rowData', allow_duplicate=True),
    Output('modal-confirma-asist', 'is_open', allow_duplicate=True),

    Input('btn-confirma-asist', 'n_clicks'),
    State('selector-confirma-asist', 'value'),
    State('parametros', 'data'),
)
def cambia_opcion_asiste(click, asiste, param):
    if click == 0:
        raise PreventUpdate
    else:
        modifica_asiste(param['usuario'], param['id_visita'], asiste)
        data = Actualiza(param)

        return data.asisten_dic(), data.texto_advertencia(), data.oculta_advertencia(), data.invitaciones(param['cond_asiste']), False  # cambio aquí


# --------------------------------------------------------------------------
# añade las passwords
@app.callback(
    Output('inp-pw', 'value'),
    Input('sel-u', 'value'),
)
def agrega_password(univ):
    return os.environ.get('U'+str(univ))


if __name__ == '__main__':
    app.run()
