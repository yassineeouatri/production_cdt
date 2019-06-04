# -*- coding: utf-8 -*-
from osv import osv ,fields
from tools.translate import _
import time
import tools
import pymssql
import logging
import pyodbc
import sys
import os
import csv
import datetime
import xlsxwriter
_logger = logging.getLogger(__name__)

AVAILABLE_JOURS =[('Lundi', 'Lundi'), ('Mardi', 'Mardi'), \
                                  ('Mercredi', 'Mercredi'), ('Jeudi', 'Jeudi'),\
                                  ('Vendredi', 'Vendredi'), ('Samedi', 'Samedi'), \
                                  ('Dimanche', 'Dimanche')]
AVAILABLE_MODES = [('nuit', 'Nuit'),('jour', 'Jour')]
AVAILABLE_RELANCES = [('oui', 'OUI'),('non', 'NON')]
AVAILABLE_STATES = [
    ('present', 'PRESENT'),
    ('absent', 'ABSENT'),
    ('retard', 'RETARD'),
    ('depassement', 'DEPASSEMENT'),
    ('en_conge','CONGE'),
    ('abandon','ABANDON'),
    ('arret','ARRET'),
    ('destaff','DESTAFF'),
]
AVAILABLE_MONTH =[(1, 'Janvier'), (2, 'Février'), \
                                  (3, 'Mars'), (4, 'Avril'),\
                                  (5, 'Mai'), (6, 'Juin'), \
                                  (7, 'Juillet'), (8, 'Août'),\
                                  (9, 'Septembre'), (10, 'Octobre'),\
                                  (11, 'Novembre'), (12, 'Décembre')]
sortie='C:\\Odoo\\addons\\web\\static\\reporting\\'
class hr_employee(osv.osv):
    _name = 'hr.employee'
    _inherit = 'hr.employee'
    _columns = {
        }
hr_employee()
class suivi_production_cdt_temp(osv.osv):
    _name = "suivi.production.cdt.temp"
    _description = "suivi.production.cdt.temp"
    _order ='c_datetime desc'
    _columns = {
        'c_base': fields.char('Base', size=64,required=True),
        'c_datetime': fields.char('Date', size=64),
        'c_call_duration' : fields.integer('Duration Call'),
        'c_agent_id': fields.integer('Agent Id'),
        'c_agent_name': fields.char('Agent Name', size=64),
        'c_agent_duration' : fields.integer('Duration Agent'),
        'c_status_id': fields.char('Status Id', size=64),
        'c_status_label': fields.char('Status', size=64),
         }
    _defaults = {} 
   
    _sql_constraints = [  ]
suivi_production_cdt_temp() 
class hr_operation_planning_cdt(osv.osv):
    _name = "hr.operation.planning.cdt"
    _description = "Planning CDT"
    _order= 'date desc'
    def _day_compute(self, cr, uid, ids, fieldnames, args, context=None):
        res = dict.fromkeys(ids, '')
        for obj in self.browse(cr, uid, ids, context=context):
            res[obj.id] = time.strftime('%d', time.strptime(obj.date, '%Y-%m-%d'))
        return res
    def _heure_prevu(self, cr, uid, ids, fieldnames, args, context=None):
        res = dict.fromkeys(ids, '')
        for obj in self.browse(cr, uid, ids, context=context):
            rslt=0
            rslt=obj.am_dep+obj.pm_dep-obj.am_arr-obj.pm_arr
        res[obj.id]=rslt
        return res
    def get_operation_id(self, cr, uid, context=None):
        operation_obj=self.pool.get('hr.operation')
        operation_ids = operation_obj.search(cr, uid, [('name','=','CDT')], context=context)
        for obj in operation_obj.browse(cr, uid, operation_ids, context=context): 
            operation_id=obj.id  
        return operation_id
    def create(self, cr, uid, data, context=None):
        planning_id = super(hr_operation_planning_cdt, self).create(cr, uid, data, context=context)
        plannings_ids = self.search(cr, uid, [('id','=',planning_id)], context=context)
        for obj in self.browse(cr, uid, plannings_ids, context=context): 
            operation_id=obj.operation_id.id
            date=obj.date
        self.pool.get('hr.employee.presence.cdt').start_calcul(cr, uid, date,operation_id,data, context)
        return planning_id
    
     
    _columns = {
        'user_id': fields.many2one('res.users', 'Utilisateur',readonly=True),
        'create_date': fields.datetime('Date',readonly=True),
        'date': fields.date('Date',required=True,readonly=False),
        'day': fields.function(_day_compute, type='char', string='Jour', store=True, select=1, size=32),
        'operation_id': fields.many2one('hr.operation', "Opération",required=True,readonly=True),
        'h_prevu':fields.function(_heure_prevu,type='float', string='H.Prévu',store=True),
        'am_arr' : fields.float("Matin Arr.",digit=(6,2)),
        'am_dep' : fields.float("Matin Dép.",digit=(6,2)),
        'pm_arr' : fields.float("AP-MIDI Arr.",digit=(6,2)),
        'pm_dep' : fields.float("AP-MIDI  Dép.",digit=(6,2)),
        's_dep' : fields.float("Soir Dép.",digit=(6,2)),
        's_arr' : fields.float("Soir Arr.",digit=(6,2))
         }
    _defaults = {
        'date': lambda *a: time.strftime('%Y-%m-%d'),
        'user_id':  lambda self, cr, uid, context: uid,
        'operation_id' : get_operation_id
         } 
   
    _sql_constraints = [
        ('name_uniq', 'unique (date,operation_id)', 'Le couple(operation,date) doit etre unique!')
    ]
hr_operation_planning_cdt()
class hr_employee_dimensionnement_cdt(osv.osv):
    _name = 'hr.employee.dimensionnement.cdt'
    _order ='date desc'
    def get_h_pres(self, cr, uid, ids, fieldnames, args, context=None):
        res = dict.fromkeys(ids, '')
        for obj in self.browse(cr, uid, ids, context=context):
            rslt=0
            if obj.presence_ids:
                for x in obj.presence_ids:
                    if x.category_id.name=='TA':
                        rslt+=x.h_pres
            res[obj.id]=rslt
        return res
    def get_h_fact(self, cr, uid, ids, fieldnames, args, context=None):
        res = dict.fromkeys(ids, '')
        for obj in self.browse(cr, uid, ids, context=context):
            rslt=0
            if obj.presence_ids:
                for x in obj.presence_ids:
                    if x.category_id.name=='TA':
                        rslt+=x.h_fact
            res[obj.id]=rslt
        return res
    def get_cu(self, cr, uid, ids, fieldnames, args, context=None):
        res = dict.fromkeys(ids, '')
        for obj in self.browse(cr, uid, ids, context=context):
            rslt=0
            if obj.production_ids:
                for x in obj.production_ids:
                    rslt+=x.cu
            res[obj.id]=rslt
        return res
    def get_h_prod(self, cr, uid, ids, fieldnames, args, context=None):
        res = dict.fromkeys(ids, '')
        for obj in self.browse(cr, uid, ids, context=context):
            rslt=0
            if obj.production_ids:
                for x in obj.production_ids:
                    rslt+=x.h_prod_jour+x.h_prod_nuit+x.h_brief
            res[obj.id]=rslt
        return res
    def get_diff(self, cr, uid, ids, fieldnames, args, context=None):
        res = dict.fromkeys(ids, '')
        for obj in self.browse(cr, uid, ids, context=context):
            presence=prod=0
            if obj.presence_ids:
                for x in obj.presence_ids:
                    if x.category_id.name=='TA':
                        presence+=x.h_fact
            if obj.production_ids:
                for x in obj.production_ids:
                    prod+=x.h_prod_jour+x.h_prod_nuit+x.h_brief
            res[obj.id]=presence-prod
        return res
    _columns = {
        'name': fields.datetime('Date Creation',readonly=False),
        'day': fields.char('Day',32,),
        'jour':fields.selection(AVAILABLE_JOURS,'Jour',readonly=False,required=True),
        'date' : fields.date('Date',readonly=False,required=True),
        'h_pres' : fields.function(get_h_pres,type='float', string='H.Présence TA'),
        'h_fact' : fields.function(get_h_fact,type='float', string='H.Facturables TA'),
        'h_prod' : fields.function(get_h_prod,type='float', string='H.Production'),
        'h_diff' : fields.function(get_diff,type='float', string='Différence'),
        'cu' : fields.function(get_cu,type='integer', string='Invitation'),
        'superviseur_id': fields.many2one('hr.employee', "Equipe", readonly=False),
              
        'presence_ids' : fields.one2many('hr.employee.presence.cdt' ,'dimensionnement_id','Présences'),
        'production_ids' : fields.one2many('suivi.production.ta.cdt' ,'dimensionnement_id','Productions')
        }
hr_employee_dimensionnement_cdt()
class hr_employee_presence_cdt(osv.osv):
    _name = 'hr.employee.presence.cdt'
    _order ='date desc,category_id,employee_id'
    def _day_compute(self, cr, uid, ids, fieldnames, args, context=None):
        res = dict.fromkeys(ids, '')
        for obj in self.browse(cr, uid, ids, context=context):
            res[obj.id] = obj.date
        return res
    def _heure_presence(self, cr, uid, ids, fieldnames, args, context=None):
        res = dict.fromkeys(ids, '')
        for obj in self.browse(cr, uid, ids, context=context):
            rslt=0
            rslt=obj.am_dep+obj.pm_dep-obj.am_arr-obj.pm_arr
        res[obj.id]=rslt
        return res
    def _heure_facturable(self, cr, uid, ids, fieldnames, args, context=None):
        res = dict.fromkeys(ids, '')
        for obj in self.browse(cr, uid, ids, context=context):
            rslt=0
            rslt=obj.am_dep+obj.pm_dep-obj.am_arr-obj.pm_arr-obj.h_brief-obj.h_panne-obj.h_formation-obj.h_relance
        res[obj.id]=rslt
        return res
    def start_calcul(self, cr, uid, date,operation_id,ids=True, context=True):
        cr.execute("select date('"+date+"'),to_char(date('"+date+"'),'d'),emp.id,emp.category_id,pl.operation_id,\
                    am_arr,am_dep,pm_arr,pm_dep,s_arr,s_dep,h_prevu,\
                    case when state='open' then 'present' else state end,superviseur_id\
                    from hr_operation_planning_cdt pl\
                    left join hr_employee emp on emp.operation_id=pl.operation_id \
                    where pl.date=date('"+date+"')\
                    and state in ('open','en_conge','arret_maladie')\
                    and emp.category_id in (select id from hr_employee_category where name in ('SUP','TA','Chef de plateau'))\
                    and to_char(date('"+date+"'),'YYYY-mm-dd')||pl.operation_id||emp.id not in\
                    (select distinct to_char(date,'YYYY-mm-dd')||operation_id||employee_id\
                    from hr_employee_presence_cdt where date=date('"+date+"'))",(tuple(),)) 
        for res in cr.fetchall():
            jour=res[1]
            if jour=='1':
                jour='Dimanche'
            if jour=='2':
                jour='Lundi'
            if jour=='3':
                jour='Mardi'
            if jour=='4':
                jour='Mercredi'
            if jour=='5':
                jour='Jeudi'
            if jour=='6':
                jour='Vendredi'
            if jour=='7':
                jour='Samedi'
            if res[12]=="en_conge" :
                presence_id = self.create(cr,uid,{
                                                'jour':jour,
                                                'date' : res[0],
                                                'employee_id': res[2],
                                                'superviseur_id' : res[13],
                                                'category_id': res[3],
                                                'operation_id': res[4],
                                                'am_arr' : 0,
                                                'am_dep' : 0,
                                                'pm_arr' : 0,
                                                'pm_dep' : 0,
                                                's_arr' : 0,
                                                's_dep' : 0,
                                                'h_prevu': 0,
                                                'state':res[12],
                                                  })
            else :
                presence_id = self.create(cr,uid,{
                                                'jour':jour,
                                                'date' : res[0],
                                                'employee_id': res[2],
                                                'superviseur_id' : res[13],
                                                'category_id': res[3],
                                                'operation_id': res[4],
                                                'am_arr' : res[5],
                                                'am_dep' : res[6],
                                                'pm_arr' : res[7],
                                                'pm_dep' : res[8],
                                                's_arr' : res[9],
                                                's_dep' : res[10],
                                                'h_prevu': res[11],
                                                'state':res[12],
                                                  })
            cr.commit()
        return True
    def start_calcul_date(self, cr, uid, date,ids=True, context=True):
        cr.execute("select date('"+date+"'),to_char(date('"+date+"'),'d'),emp.id,emp.category_id,pl.operation_id,\
                    am_arr,am_dep,pm_arr,pm_dep,s_arr,s_dep,h_prevu,\
                    case when state='open' then 'present' else state end,superviseur_id\
                    from hr_operation_planning_cdt pl\
                    left join hr_employee emp on emp.operation_id=pl.operation_id \
                    where pl.date=date('"+date+"')\
                    and state in ('open','en_conge','arret_maladie')\
                    and emp.category_id in (select id from hr_employee_category where name in ('SUP','TA','Chef de plateau'))\
                    and to_char(date('"+date+"'),'YYYY-mm-dd')||pl.operation_id||emp.id not in\
                    (select distinct to_char(date,'YYYY-mm-dd')||operation_id||employee_id\
                    from hr_employee_presence_cdt where date=date('"+date+"'))\
                     and date('"+date+"') >= date(date_entree)",(tuple(),)) 
        for res in cr.fetchall():
            jour=res[1]
            if jour=='1':
                jour='Dimanche'
            if jour=='2':
                jour='Lundi'
            if jour=='3':
                jour='Mardi'
            if jour=='4':
                jour='Mercredi'
            if jour=='5':
                jour='Jeudi'
            if jour=='6':
                jour='Vendredi'
            if jour=='7':
                jour='Samedi'
            if res[12]=="en_conge" :
                presence_id = self.create(cr,uid,{
                                                'jour':jour,
                                                'date' : res[0],
                                                'employee_id': res[2],
                                                'superviseur_id' : res[13],
                                                'category_id': res[3],
                                                'operation_id': res[4],
                                                'am_arr' : 0,
                                                'am_dep' : 0,
                                                'pm_arr' : 0,
                                                'pm_dep' : 0,
                                                's_arr' : 0,
                                                's_dep' : 0,
                                                'h_prevu': 0,
                                                'state':res[12],
                                                  })
            else :
                presence_id = self.create(cr,uid,{
                                                'jour':jour,
                                                'date' : res[0],
                                                'employee_id': res[2],
                                                'superviseur_id' : res[13],
                                                'category_id': res[3],
                                                'operation_id': res[4],
                                                'am_arr' : res[5],
                                                'am_dep' : res[6],
                                                'pm_arr' : res[7],
                                                'pm_dep' : res[8],
                                                's_arr' : res[9],
                                                's_dep' : res[10],
                                                'h_prevu': res[11],
                                                'state':res[12],
                                                  })
            cr.commit()
        return True
    def create_hr_employee_presence(self, cr, uid, ids=True, context=True):
        cr.execute("""select date(now()),to_char(date(now()),'d'),emp.id,emp.category_id,pl.operation_id,am_arr,am_dep,pm_arr,pm_dep,s_arr,s_dep,h_prevu,
                    case when state='open' then 'present' else state end,superviseur_id
                    from hr_operation_planning_cdt pl
                    left join hr_employee emp on emp.operation_id=pl.operation_id 
                    where pl.date=date(now())
                    and state in ('open','en_conge','arret_maladie')
                    and emp.category_id in (select id from hr_employee_category where name in ('SUP','TA','Chef de plateau'))
                    and to_char(date(now()),'YYYY-mm-dd')||pl.operation_id||emp.id not in
                    (select distinct to_char(date,'YYYY-mm-dd')||operation_id||employee_id from hr_employee_presence_cdt where date=date(now()))""",(tuple(),)) 
        for res in cr.fetchall():
            jour=res[1]
            if jour=='1':
                jour='Dimanche'
            if jour=='2':
                jour='Lundi'
            if jour=='3':
                jour='Mardi'
            if jour=='4':
                jour='Mercredi'
            if jour=='5':
                jour='Jeudi'
            if jour=='6':
                jour='Vendredi'
            if jour=='7':
                jour='Samedi'
            if res[12]=="en_conge" :
                presence_id = self.create(cr,uid,{
                                                'jour':jour,
                                                'date' : res[0],
                                                'employee_id': res[2],
                                                'superviseur_id': res[13],
                                                'category_id': res[3],
                                                'operation_id': res[4],
                                                'am_arr' : 0,
                                                'am_dep' : 0,
                                                'pm_arr' : 0,
                                                'pm_dep' : 0,
                                                's_arr' : 0,
                                                's_dep' : 0,
                                                'h_prevu': 0,
                                                'state':res[12],
                                                  })
            else :
                presence_id = self.create(cr,uid,{
                                                'jour':jour,
                                                'date' : res[0],
                                                'employee_id': res[2],
                                                'superviseur_id': res[13],
                                                'category_id': res[3],
                                                'operation_id': res[4],
                                                'am_arr' : res[5],
                                                'am_dep' : res[6],
                                                'pm_arr' : res[7],
                                                'pm_dep' : res[8],
                                                's_arr' : res[9],
                                                's_dep' : res[10],
                                                'h_prevu': res[11],
                                                'state':res[12],
                                                  })
            cr.commit()
        return True
    def get_h_prod(self, cr, uid, ids, fieldnames, args, context=None):
        res = dict.fromkeys(ids, '')
        for obj in self.browse(cr, uid, ids, context=context):
            rslt=0
            if obj.production_ids:
                for x in obj.production_ids:
                    rslt+=x.h_prod_jour+x.h_prod_nuit+x.h_brief
            res[obj.id]=rslt
        return res
    def get_diff(self, cr, uid, ids, fieldnames, args, context=None):
        res = dict.fromkeys(ids, '')
        for obj in self.browse(cr, uid, ids, context=context):
            presence=prod=0
            presence+=obj.h_fact
            if obj.production_ids:
                for x in obj.production_ids:
                    prod+=x.h_prod_jour+x.h_prod_nuit+x.h_brief
            res[obj.id]=presence-prod
        return res
    def get_nom_colonne(self,numero):
        chif1=chif2=chif3=''
        num=(ord('A')+numero-65)
        num1 = num / 676
        num-=676*num1
        num2=num/26
        num-= 26*num2
        if(num1>0):
            chif1=chr(num+64)
        if(num2>0):
            chif2=chr(num2+64)
        chif3=chr(num+65)
        character=chif1+chif2+chif3
        return character
    def get_nom_jour(self,jour_num):
        jour=''
        if (jour_num=='1'):
            jour='Dimanche'
        elif (jour_num=='2'):
            jour='Lundi'
        elif (jour_num=='3'):
            jour='Mardi'
        elif (jour_num=='4'):
            jour='Mercredi'
        elif (jour_num=='5'):
            jour='Jeudi'
        elif (jour_num=='6'):
            jour='Vendredi'
        elif (jour_num=='7'):
            jour='Samedi'
        return jour
    def get_nom_mois(self,jour_num):
        jour=''
        if (jour_num=='01'):
            jour='Janvier'
        elif (jour_num=='02'):
            jour='Février'
        elif (jour_num=='03'):
            jour='Mars'
        elif (jour_num=='04'):
            jour='Avril'
        elif (jour_num=='05'):
            jour='Mai'
        elif (jour_num=='06'):
            jour='Juin'
        elif (jour_num=='07'):
            jour='Juillet'
        elif (jour_num=='08'):
            jour='Août'
        elif (jour_num=='09'):
            jour='Septembre'
        elif (jour_num=='10'):
            jour='Octobre'
        elif (jour_num=='11'):
            jour='Novembre'
        elif (jour_num=='12'):
            jour='Décembre'
        return jour
    
    
    _columns = {
        'user_id': fields.many2one('res.users', 'Utilisateur',readonly=False),
        'name': fields.datetime('Date Creation',readonly=False),
        'jour':fields.selection(AVAILABLE_JOURS,'Jour',readonly=False),
        'date' : fields.date('Date',readonly=False),
        
        'day': fields.function(_day_compute, type='char', string='Day', store=True, select=1, size=32,readonly=False),
        'employee_id': fields.many2one('hr.employee', "Nom et Prénom", readonly=False),
        'superviseur_id': fields.many2one('hr.employee', "Equipe", readonly=False),
        'operation_id': fields.many2one('hr.operation', "Opération", readonly=False),
        'category_id': fields.many2one('hr.employee.category', "Catégorie", readonly=False ),
        
        'am_arr' : fields.float("Matin Arr.",digit=(6,2)),
        'am_dep' : fields.float("Matin Dép.",digit=(6,2)),
        'pm_arr' : fields.float("AP-MIDI Arr.",digit=(6,2)),
        'pm_dep' : fields.float("AP-MIDI  Dép.",digit=(6,2)),
        's_dep' : fields.float("Soir Dép.",digit=(6,2)),
        's_arr' : fields.float("Soir Arr.",digit=(6,2)),
        
        'h_prevu': fields.float("H.Prevu",digit=(6,2)),
        
        'h_pres':fields.function(_heure_presence,type='float', string='H.Présence',store=True),
        'h_prod' : fields.function(get_h_prod,type='float', string='H.Production'),
        'h_diff' : fields.function(get_diff,type='float', string='Différence'),
         
        'h_brief':fields.float("Brief Telexcel",digit=(6,2)),
        'h_panne':fields.float("Panne",digit=(6,2)),
        'h_formation':fields.float("Formation Initiale",digit=(6,2)),
        'h_relance':fields.float("H.Relance Non Facturée",digit=(6,2)),
        'h_relance_facture':fields.float("H.Relance Facturée",digit=(6,2)),
        
        'h_fact':fields.function(_heure_facturable,type='float', string='H.Fact',store=True),
        'dimensionnement_id' : fields.many2one('hr.employee.dimensionnement.cdt' ,' Dimensionnement'),
        'state': fields.selection(AVAILABLE_STATES, 'statut', size=64),
        'production_ids' : fields.one2many('suivi.production.ta.cdt' ,'presence_id','Productions')
        
        
        }
    def onchange_state(self, cr, uid, ids, state, context=None):
        value={}
        if state=='abandon' or state=='arret' or state=='en_conge':
            for obj in self.browse(cr, uid, ids, context=context):
                cr.execute("update hr_employee set state='"+state+"' \
                            where id="+str(obj.employee_id.id),(tuple(ids),))
                cr.commit()
            if state=='en_conge' :
                    self.write(cr, uid, ids, {'am_arr': 0,'am_dep': 0,'pm_arr': 0,'pm_dep': 0,'s_dep':0,'s_arr':0,
                                              'h_prevu': 0,'h_pres':0,
                                      'h_brief': 0,'h_panne': 0,'h_formation': 0,'h_relance': 0,'h_fact':0})
        if state =='absent' or state=='present':
            for obj in self.browse(cr, uid, ids, context=context):
                cr.execute("update hr_employee set state='open' \
                            where id="+str(obj.employee_id.id),(tuple(ids),))
                cr.commit()
                if state=='absent' :
                    self.write(cr, uid, ids, {'am_arr': 0,'am_dep': 0,'pm_arr': 0,'pm_dep': 0,'s_dep':0,'s_arr':0,
                                              'h_prevu': 0,'h_pres':0,
                                      'h_brief': 0,'h_panne': 0,'h_formation': 0,'h_relance': 0,'h_fact':0})
        return {'value': value}
hr_employee_presence_cdt()
class suivi_production_ta_cdt(osv.osv):
    _name = 'suivi.production.ta.cdt'
    _order ='date desc,tv'
    
    def get_data_hermes(self, cr, uid,date, ids=True, context=None):  
        suivi_production_cdt_temp=self.pool.get('suivi.production.cdt.temp')      
        cr.execute("delete  from suivi_production_cdt_temp")
        cr.commit()
        cr.execute("select base from production_base_work_cdt where date='"+str(date)+"' order by base",(tuple(),))
        for res in cr.fetchall():
            base= res[0]
            _logger.info(base)
            cr.execute("select server_host,server_user,server_password from production_server_cdt",(tuple(),))
            for res in cr.fetchall():
                server_host=res[0]
                server_user=res[1]
                server_password=res[2]
                try :
                    _logger.info(base)
                    _logger.info(server_host)
                    conn = pymssql.connect(host=server_host,user=server_user,password=server_password,database=base)
                    cur = conn.cursor()  
                    query="select tv,lib_status,duree,id_tv  from "+base+".dbo.appels\
                             where duree is not null\
                             and date="+str(date)
                    _logger.info(query)
                
                    cur.execute(query)
                    for res2 in cur:
                        tv=res2[0]
                        lib_status=res2[1]
                        duree = res2[2]
                        id_tv = res2[3]
                        suivi_production_cdt_temp.create(cr,uid,{
                                'c_base': base,
                                'c_datetime': date,
                                'c_call_duration' : duree,
                                'c_agent_id': id_tv,
                                'c_agent_name': tv,
                                'c_agent_duration' : duree,
                                'c_status_id': lib_status,
                                'c_status_label': lib_status, 
                            })
                        cr.commit()
                    conn.close()
                except :
                    pass
           
        return True
    def get_data(self, cr, uid, date, ids=True, context=None):
        reload(sys)
        sys.setdefaultencoding("UTF8")
        suivi_production_cdt_temp=self.pool.get('suivi.production.cdt.temp')
        path="C:\\Documents and Settings\\Administrateur\\Bureau\\CDT\\"+str(date)
        if os.path.exists(path):
            for path, dirs, files in os.walk(path):
                for filename in files:
                    base=os.path.splitext(filename)[0]
                    _logger.info(base)
                    ifile  = open(path+'\\'+filename, "rb")
                    reader = csv.reader(ifile, delimiter =';')
                    rownum = 0
                    for row in reader:
                        # Save header row.
                        if rownum == 0:
                            header = row
                        else:
                            
                            colnum = 0
                            c_system_status=''
                            for col in row:
                                if header[colnum]=='c_call_duration':
                                    c_call_duration=col 
                                if header[colnum]=='c_datetime':
                                    c_datetime=col 
                                if header[colnum]=='c_status_id':
                                    c_status_id=col
                                if header[colnum]=='c_agent_name':
                                    c_agent_name=col
                                if header[colnum]=='c_agent_id':
                                    c_agent_id=col
                                if header[colnum]=='c_agent_duration':
                                    c_agent_duration=col
                                if header[colnum]=='c_status_label':
                                    c_status_label=col
                                colnum += 1
                           
                            suivi_production_cdt_temp.create(cr,uid,{
                                    'c_base': base,
                                    'c_system_status': c_system_status,
                                    'c_datetime': c_datetime,
                                    'c_call_duration' : c_call_duration,
                                    'c_agent_id': c_agent_id,
                                    'c_agent_name': c_agent_name,
                                    'c_agent_duration' : c_agent_duration,
                                    'c_status_id': c_status_id,
                                    'c_status_label': c_status_label, 
                                })
                            cr.commit()
                        rownum += 1
        else :
            os.mkdir(path)
           
        
        return True
    def update_data(self, cr, uid, date, ids=True, context=None):
        cr.execute("update suivi_production_cdt_temp set c_agent_name=trim(c_agent_name) where c_agent_name like '% ';")
        cr.commit()
        base_object = self.pool.get('production.base.cdt')
        query="select distinct replace(replace(replace(replace(replace(c_base,'1',''),'2',''),'3',''),'4',''),'5','')\
                from suivi_production_cdt_temp where date(c_datetime)=date('"+date+"') \
                and c_datetime not like ''"
        cr.execute(query)
        for res in cr.fetchall():
            base=res[0]
            
            cr.execute("select count(*)  from production_base_cdt where name='"+base+"'",(tuple(),))
            #raise osv.except_osv(_("select count(*)  from production_base_cdt where name='"+base+"'",))
            for res1 in cr.fetchall():
                if res1[0]==0:
                    base_id=base_object.create(cr,uid,{  'name' : base})
                    cr.commit()  
            cr.execute("select to_char(date('"+str(date)+"') ,'D')",(tuple(),))
            for res1 in cr.fetchall():
                jour=res1[0]
            
            ################Récupérer le jour#######################
            if jour=='2':
                jour='Lundi'
            if jour=='3':
                jour='Mardi'
            if jour=='4':
                jour='Mercredi'
            if jour=='5':
                jour='Jeudi'
            if jour=='6':
                jour='Vendredi'
            if jour=='7':
                jour='Samedi'
            if jour=='1':
                jour='Dimanche'                   
            ##################    Récupérer base_id###################
            cr.execute("select id,is_relance  from production_base_cdt where name='"+base+"'",(tuple(),))
            for res1 in cr.fetchall():
                base_id=res1[0]
                is_relance=res1[1]
            if is_relance is not None:
                if is_relance=='oui':
                ################        Récupérer les données ###############
                    query1="select c_agent_name,0\
                            from suivi_production_cdt_temp \
                            where replace(replace(replace(replace(replace(c_base,'1',''),'2',''),'3',''),'4',''),'5','')='"+base+"'\
                            and date(c_datetime)='"+str(date)+"'\
                            and c_agent_name is not null\
                             and c_datetime not like ''\
                            group by c_agent_name"
                else :
                    query1="select c_agent_name,count(case when c_status_label like 'OK%' then c_status_label end)\
                            from suivi_production_cdt_temp \
                            where replace(replace(replace(replace(replace(c_base,'1',''),'2',''),'3',''),'4',''),'5','')='"+base+"'\
                            and date(c_datetime)='"+str(date)+"'\
                             and c_datetime not like ''\
                             and c_agent_name is not null\
                            group by c_agent_name\
                           "
                _logger.info(query1)
                cr.execute(query1)
                for res2 in cr.fetchall():
                    tv=res2[0]
                    cu=res2[1]
                    ################################################################
                    cr.execute("select count(*) from suivi_production_ta_cdt \
                        where tv='"+tv.replace("'","''")+"'\
                        and base_id='"+str(base_id)+"'\
                        and date ='"+str(date)+"'",tuple())
                    for res_test in cr.fetchall():
                        if res_test[0]==0:
                            self.create(cr,uid,{
                                                'jour':jour,
                                                'date' : date,
                                                'base_id' : base_id,
                                                'tv' : tv,
                                                'cu' : cu,
                                                    })
                            cr.commit()
                        else :
                            cr.execute("update suivi_production_ta_cdt \
                                    set cu="+str(cu)+",\
                                    cadence1=case when (case when h_prod_jour is null then 0 else h_prod_jour end+case when h_prod_nuit is null then 0 else h_prod_nuit end+case when h_brief is null then 0 else h_brief end+case when h_panne is null then 0 else h_panne end+case when h_relance is null then 0 else h_relance end)=0 then 0 else round(cast("+str(cu)+"*1.00/(case when h_prod_jour is null then 0 else h_prod_jour end+case when h_prod_nuit is null then 0 else h_prod_nuit end+case when h_brief is null then 0 else h_brief end+case when h_panne is null then 0 else h_panne end+case when h_relance is null then 0 else h_relance end) as numeric),2) end,\
                                    cadence2=case when (case when h_prod_jour is null then 0 else h_prod_jour end+case when h_prod_nuit is null then 0 else h_prod_nuit end+case when h_brief is null then 0 else h_brief end+case when h_relance is null then 0 else h_relance end)=0 then 0 else round(cast("+str(cu)+"*1.00/(case when h_prod_jour is null then 0 else h_prod_jour end+case when h_prod_nuit is null then 0 else h_prod_nuit end+case when h_brief is null then 0 else h_brief end+case when h_relance is null then 0 else h_relance end) as numeric),2) end,\
                                    cadence3=case when (case when h_prod_jour is null then 0 else h_prod_jour end+case when h_prod_nuit is null then 0 else h_prod_nuit end+case when h_brief is null then 0 else h_brief end)=0 then 0 else round(cast("+str(cu)+"*1.00/(case when h_prod_jour is null then 0 else h_prod_jour end+case when h_prod_nuit is null then 0 else h_prod_nuit end+case when h_brief is null then 0 else h_brief end) as numeric),2) end\
                                    where tv='"+tv.replace("'","''")+"'\
                                    and base_id='"+str(base_id)+"'\
                                    and date ='"+str(date)+"'",tuple())
                            cr.commit()
        
        ################################################
        self.pool.get('hr.employee.presence.cdt').start_calcul_date(cr, uid, date, context)
       
        self.pool.get('hr.employee.presence.cdt').create_hr_employee_presence( cr, uid, id, context)
        cr.execute("""
            select distinct date,day,jour,superviseur_id
                 from hr_employee_presence_cdt
                 where date::text||superviseur_id::text not in(select date::text||superviseur_id::text from hr_employee_dimensionnement_cdt where superviseur_id is not null)
                  and superviseur_id is not null
                  and date>=date(now())-5""",(tuple(),)) 
        for res in cr.fetchall():
            self.pool.get('hr.employee.dimensionnement.cdt').create(cr,uid,{  'date' : res[0],
                                                                                'day' : res[1],
                                                                                'jour' : res[2],
                                                                                'superviseur_id' : res[3]})
        ######################## Mis à jour dimensionnementd_id########################
        cr.execute("""select pres.id,dim.id
                        from hr_employee_presence_cdt pres
                        inner join hr_employee_dimensionnement_cdt dim 
                        on dim.date=pres.date and dim.superviseur_id=pres.superviseur_id
                        where pres.dimensionnement_id is null""",(tuple(),)) 
        for res in cr.fetchall():
            cr.execute("update hr_employee_presence_cdt set dimensionnement_id="+str(res[1])+" where id="+str(res[0])+"",(tuple(),))   
            cr.commit()
        ######################## Mis à jour dimensionnementd_id########################
        cr.execute("""select prod.id,dim.id
                        from suivi_production_ta_cdt prod
                        inner join hr_employee_dimensionnement_cdt dim 
                        on dim.date=prod.date and dim.superviseur_id=prod.superviseur_id
                        where prod.dimensionnement_id is null
                        and dim.id is not null""",(tuple(),)) 
        for res in cr.fetchall():
            cr.execute("update suivi_production_ta_cdt set dimensionnement_id="+str(res[1])+" where id="+str(res[0])+"",(tuple(),))   
            cr.commit()
        ######################## Mis à jour presence_id########################
        cr.execute("""select prod.id,pres.id 
                        from suivi_production_ta_cdt prod
                        left join hr_employee_presence_cdt pres
                        on pres.date=prod.date and pres.employee_id=prod.employee_id
                        where pres.id is not null and prod.presence_id is null""",(tuple(),)) 
        for res in cr.fetchall():
            cr.execute("update suivi_production_ta_cdt set presence_id="+str(res[1])+" where id="+str(res[0])+"",(tuple(),))   
            cr.commit()
        ##################### Mettre à jour la liste des logins ###################
        cr.execute("""select distinct tv from suivi_production_ta_cdt
            where tv is not null
            and tv not in(select name from production_login_cdt where name is not null)""",tuple())
        for res in cr.fetchall():
            self.pool.get('production.login.cdt').create(cr,uid,{'name':res[0],})
            cr.commit()
        ################################################
        cr.execute("""select prod.id,emp.id,emp.superviseur_id
                         from suivi_production_ta_cdt prod 
                         left join production_login_cdt log on log.name=prod.tv
                         left join hr_employee emp on emp.id=log.employee_id
                         where emp.id is not null
                         and prod.employee_id is null
                         and emp.superviseur_id is not null """,tuple())
        for res in cr.fetchall():
            cr.execute("update suivi_production_ta_cdt set employee_id="+str(res[1])+"\
                        ,superviseur_id="+str(res[2])+"\
                        where id="+str(res[0]),tuple())
            cr.commit()
        ##########################################################
        cr.execute("""delete from production_update_cdt  where objet is null""",tuple())
        cr.commit()
        cr.execute("update production_update_cdt set date_update=(\
                    select max(date) from suivi_production_ta_cdt) where objet='suivi_production_ta_cdt'",(tuple(),))
        cr.commit()
        return True
    
    
    def _day_compute(self, cr, uid, ids, fieldnames, args, context=None):
        res = dict.fromkeys(ids, '')
        for obj in self.browse(cr, uid, ids, context=context):
            res[obj.id] = time.strftime('%Y-%m-%d', time.strptime(obj.date, '%Y-%m-%d'))
        return res
    def get_tx_brief(self, cr, uid, ids, fieldnames, args, context=None):
        res = dict.fromkeys(ids, '')
        for obj in self.browse(cr, uid, ids, context=context):
            rslt=0.00
            heure_total=obj.h_prod_jour+obj.h_prod_nuit+obj.h_brief+obj.h_panne+obj.h_relance
            if heure_total!=0 :
                rslt=obj.h_brief*100.00/heure_total
        res[obj.id]=str(round(rslt,2))+' %'
        return res
    def get_tx_panne(self, cr, uid, ids, fieldnames, args, context=None):
        res = dict.fromkeys(ids, '')
        for obj in self.browse(cr, uid, ids, context=context):
            rslt=0.00
            heure_total=obj.h_prod_jour+obj.h_prod_nuit+obj.h_brief+obj.h_panne+obj.h_relance
            if heure_total!=0 :
                rslt=obj.h_panne*100.00/heure_total
        res[obj.id]=str(round(rslt,2))+' %'
        return res
    def get_total_heure(self, cr, uid, ids, fieldnames, args, context=None):
        res = dict.fromkeys(ids, '')
        for obj in self.browse(cr, uid, ids, context=context):
            heure_total=obj.h_prod_jour+obj.h_prod_nuit+obj.h_brief+obj.h_panne+obj.h_relance
        res[obj.id]=heure_total
        return res
    def get_cadence1(self, cr, uid, ids, fieldnames, args, context=None):
        res = dict.fromkeys(ids, '')
        for obj in self.browse(cr, uid, ids, context=context):
            rslt=0.00
            heure_total=obj.h_prod_jour+obj.h_prod_nuit+obj.h_brief+obj.h_panne+obj.h_relance
            if heure_total!=0 :
                rslt=obj.cu*1.00/heure_total
        res[obj.id]=str(round(rslt,2))
        return res
    def get_cadence2(self, cr, uid, ids, fieldnames, args, context=None):
        res = dict.fromkeys(ids, '')
        for obj in self.browse(cr, uid, ids, context=context):
            rslt=0.00
            heure_total=obj.h_prod_jour+obj.h_prod_nuit+obj.h_brief+obj.h_relance
            if heure_total!=0 :
                rslt=obj.cu*1.00/heure_total
        res[obj.id]=str(round(rslt,2))
        return res
    def get_cadence3(self, cr, uid, ids, fieldnames, args, context=None):
        res = dict.fromkeys(ids, '')
        for obj in self.browse(cr, uid, ids, context=context):
            rslt=0.00
            heure_total=obj.h_prod_jour+obj.h_prod_nuit+obj.h_brief
            if heure_total!=0 :
                rslt=obj.cu*1.00/heure_total
        res[obj.id]=str(round(rslt,2))
        return res
    def get_statut(self, cr, uid, ids, fieldnames, args, context=None):
        res = dict.fromkeys(ids, '')
        for obj in self.browse(cr, uid, ids, context=context):
            rslt=''
            heure_facturable1=heure_facturable2=0
            if obj.employee_id.id!=False:
                cr.execute("select sum(case when h_prod_jour is null then 0 else h_prod_jour end)+sum(case when h_brief is null then 0 else h_brief end)\
                            from suivi_production_ta_cdt\
                            where  employee_id="+str(obj.employee_id.id)+"\
                            and date='"+obj.date+"'",(tuple(),))
                for res1 in cr.fetchall():
                    if res1[0]:
                        heure_facturable1=res1[0]
                cr.execute("select \
                            case when h_fact is null then 0 else h_fact end\
                            from hr_employee_presence_cdt\
                            where  employee_id="+str(obj.employee_id.id)+"\
                            and date='"+obj.date+"'",(tuple(),))
                for res1 in cr.fetchall():
                    if res1[0]:
                        heure_facturable2=res1[0]
            if heure_facturable1==heure_facturable2:
                rslt='ok'
            else :
                rslt='ko'
            res[obj.id]=rslt
        return res
    
    _columns = {
        'user_id': fields.many2one('res.users', 'Utilisateur',readonly=False),
        'jour':fields.selection(AVAILABLE_JOURS,'Jour',readonly=False),
        'date' : fields.date('Date',readonly=False),
        'day': fields.function(_day_compute, type='char', string='Day', store=True, select=1, size=32,readonly=False),
        'base_id' : fields.many2one('production.base.cdt', "Base", readonly=False),
        'employee_id': fields.many2one('hr.employee', "Nom et Prenom", readonly=False),
        'superviseur_id': fields.many2one('hr.employee', "Equipe",domain=[('state','=','open'),('operation_id.name','=','CDT'),('category_id.name','in',('SUP','Chef de plateau'))], readonly=False),
        'tv' : fields.char('TV',128,readonly=False),
        'cu' : fields.integer('CU',readonly=False),
        'h_prod_jour':fields.float("H.PROD",digit=(6,2)),
        'h_prod_nuit':fields.float("HEURES DE PROD NUIT",digit=(6,2)),
        'h_brief':fields.float("H.BRIEF",digit=(6,2)),
        'h_panne':fields.float("H.PANNE",digit=(6,2)),
        'h_relance':fields.float("H.RELANCE",digit=(6,2)),
        'h_formation':fields.float("H.FORMATION",digit=(6,2)),
        'tx_brief' : fields.function(get_tx_brief,type='char', string='TX BRIEF',size=56,store=True),
        'tx_panne' : fields.function(get_tx_panne,type='char', string='TX PANNE',size=56,store=True),
        'total_heure':fields.function(get_total_heure,type='float', string='TOTAL HEURES',store=True),
        'cadence1':fields.function(get_cadence1,type='float', string='CADENCE PANNE INCLUSES',store=True),
        'cadence2':fields.function(get_cadence2,type='float', string='CADENCE HORS PANNE',store=True),
        'cadence3':fields.function(get_cadence3,type='float', string='CADENCE HORS RELANCE ET PANNES',store=True),
        'dimensionnement_id' : fields.many2one('hr.employee.dimensionnement.cdt' ,' Dimensionnement'),
        'presence_id' : fields.many2one('hr.employee.presence.cdt' ,' Presence'),
        'statut' : fields.function(get_statut,type='char',size=16, string='Statut'),
        }
    _defaults = {
        'date': lambda *a: time.strftime('%Y-%m-%d'),
        'user_id':  lambda self, cr, uid, context: uid,
         } 
suivi_production_ta_cdt()
class suivi_production_base_cdt(osv.osv):    
    _name = "suivi.production.base.cdt"
    _description = "Suivi de production par base"
    _auto = False
    _order ='date desc, base_id asc'
   
    _columns = {
        'jour':fields.selection(AVAILABLE_JOURS,'Jour',readonly=True),
        'date' : fields.date('Date',readonly=True),
        'day': fields.char('Day',32,),
        'base_id' : fields.many2one('production.base.cdt', "Base", readonly=False),
        'cu' : fields.integer('CU'),
        'h_prod_jour':fields.float("HEURES DE PROD JOUR",digit=(6,2)),
        'h_prod_nuit':fields.float("HEURES DE PROD NUIT",digit=(6,2)),
        'h_brief':fields.float("HEURES DE BRIEF",digit=(6,2)),
        'h_panne':fields.float("HEURES DE PANNE",digit=(6,2)),
        'h_relance':fields.float("HEURES DE RELANCE",digit=(6,2)),
        'tx_brief' : fields.char('TAUX DE BRIEF',56),
        'tx_panne' : fields.char('TAUX DE PANNE',56),
        'total_heure':fields.float('TOTAL HEURES'),
        'cadence1':fields.float('CADENCE PANNE INCLUSES'),
        'cadence2':fields.float('CADENCE HORS PANNE'),
        'cadence3':fields.float('CADENCE HORS RELANCE ET PANNES'),        
        'h_formation':fields.float("HEURES DE FORMATIONS",digit=(6,2)),
        
       }
    def init(self, cr):
        tools.drop_view_if_exists(cr, 'suivi_production_base_cdt')
        cr.execute("""
            create or replace view suivi_production_base_cdt as (
                select t.id,t.jour,t.date,t.day,t.base_id,t.cu,t.h_prod_jour,t.h_prod_nuit,t.h_brief,t.h_panne,t.h_relance,t.h_formation,
                case when (t.h_prod_jour+t.h_prod_nuit+t.h_brief+t.h_panne+t.h_relance)=0 then '0.00%' else 
                round(cast(t.h_brief*100.00/(t.h_prod_jour+t.h_prod_nuit+t.h_brief+t.h_panne+t.h_relance) as numeric),2)||'%' end as tx_brief,
                case when (t.h_prod_jour+t.h_prod_nuit+t.h_brief+t.h_panne+t.h_relance)=0 then '0.00%' else 
                round(cast(t.h_panne*100.00/(t.h_prod_jour+t.h_prod_nuit+t.h_brief+t.h_panne+t.h_relance) as numeric),2)||'%' end as tx_panne,
                (t.h_prod_jour+t.h_prod_nuit+t.h_brief+t.h_panne+t.h_relance) as total_heure,
                case when (t.h_prod_jour+t.h_prod_nuit+t.h_brief+t.h_panne+t.h_relance)=0 then 0.00 else 
                round(cast(t.cu*1.00/(t.h_prod_jour+t.h_prod_nuit+t.h_brief+t.h_panne+t.h_relance) as numeric),2) end as cadence1,
                case when (t.h_prod_jour+t.h_prod_nuit+t.h_brief+t.h_relance)=0 then 0.00 else 
                round(cast(t.cu*1.00/(t.h_prod_jour+t.h_prod_nuit+t.h_brief+t.h_relance) as numeric),2) end as cadence2,
                case when (t.h_prod_jour+t.h_prod_nuit+t.h_brief)=0 then 0.00 else 
                round(cast(t.cu*1.00/(t.h_prod_jour+t.h_prod_nuit+t.h_brief) as numeric),2) end as cadence3
                from (select min(id) as id,jour,date,day,base_id,sum(cu) as cu,
                sum(case when h_prod_jour is null then 0.00 else h_prod_jour end) as h_prod_jour,
                sum(case when h_prod_nuit is null then 0.00 else h_prod_nuit end) as h_prod_nuit,
                sum(case when h_brief is null then 0.00 else h_brief end) as h_brief,
                sum(case when h_panne is null then 0.00 else h_panne end) as h_panne,
                sum(case when h_relance is null then 0.00 else h_relance end) as h_relance,
                sum(case when h_formation is null then 0.00 else h_formation end) as h_formation
                 from suivi_production_ta_cdt
                 group by jour,date,day,base_id)as t
                )
        """)
suivi_production_base_cdt()
class suivi_production_superviseur_cdt(osv.osv):    
    _name = "suivi.production.superviseur.cdt"
    _description = "Suivi de production par superviseur"
    _auto = False
    _order ='date desc, superviseur_id asc'
   
    _columns = {
        'jour':fields.selection(AVAILABLE_JOURS,'Jour',readonly=True),
        'date' : fields.date('Date',readonly=True),
        'day': fields.char('Day',32,),
        'superviseur_id' : fields.many2one('hr.employee', "SUP", readonly=False),
        'cu' : fields.integer('CU'),
        'h_prod_jour':fields.float("HEURES DE PROD JOUR",digit=(6,2)),
        'h_prod_nuit':fields.float("HEURES DE PROD NUIT",digit=(6,2)),
        'h_brief':fields.float("HEURES DE BRIEF",digit=(6,2)),
        'h_panne':fields.float("HEURES DE PANNE",digit=(6,2)),
        'h_relance':fields.float("HEURES DE RELANCE",digit=(6,2)),
        'tx_brief' : fields.char('TAUX DE BRIEF',56),
        'tx_panne' : fields.char('TAUX DE PANNE',56),
        'total_heure':fields.float('TOTAL HEURES'),
        'cadence1':fields.float('CADENCE PANNE INCLUSES'),
        'cadence2':fields.float('CADENCE HORS PANNE'),
        'cadence3':fields.float('CADENCE HORS RELANCE ET PANNES'),  
        'h_formation':fields.float("HEURES DE FORMATIONS",digit=(6,2)),      
       }
    def init(self, cr):
        tools.drop_view_if_exists(cr, 'suivi_production_superviseur_cdt')
        cr.execute("""
            create or replace view suivi_production_superviseur_cdt as (
                  select t.id,t.jour,t.date,t.day,t.superviseur_id,t.cu,t.h_prod_jour,t.h_prod_nuit,t.h_brief,t.h_panne,t.h_relance,t.h_formation,
                case when (t.h_prod_jour+t.h_prod_nuit+t.h_brief+t.h_panne+t.h_relance)=0 then '0.00%' else 
                round(cast(t.h_brief*100.00/(t.h_prod_jour+t.h_prod_nuit+t.h_brief+t.h_panne+t.h_relance) as numeric),2)||'%' end as tx_brief,
                case when (t.h_prod_jour+t.h_prod_nuit+t.h_brief+t.h_panne+t.h_relance)=0 then '0.00%' else 
                round(cast(t.h_panne*100.00/(t.h_prod_jour+t.h_prod_nuit+t.h_brief+t.h_panne+t.h_relance) as numeric),2)||'%' end as tx_panne,
                (t.h_prod_jour+t.h_prod_nuit+t.h_brief+t.h_panne+t.h_relance) as total_heure,
                case when (t.h_prod_jour+t.h_prod_nuit+t.h_brief+t.h_panne+t.h_relance)=0 then 0.00 else 
                round(cast(t.cu*1.00/(t.h_prod_jour+t.h_prod_nuit+t.h_brief+t.h_panne+t.h_relance) as numeric),2) end as cadence1,
                case when (t.h_prod_jour+t.h_prod_nuit+t.h_brief+t.h_relance)=0 then 0.00 else 
                round(cast(t.cu*1.00/(t.h_prod_jour+t.h_prod_nuit+t.h_brief+t.h_relance) as numeric),2) end as cadence2,
                case when (t.h_prod_jour+t.h_prod_nuit+t.h_brief)=0 then 0.00 else 
                round(cast(t.cu*1.00/(t.h_prod_jour+t.h_prod_nuit+t.h_brief) as numeric),2) end as cadence3
                from (select min(id) as id,jour,date,day,superviseur_id,sum(cu) as cu,
                sum(case when h_prod_jour is null then 0.00 else h_prod_jour end) as h_prod_jour,
                sum(case when h_prod_nuit is null then 0.00 else h_prod_nuit end) as h_prod_nuit,
                sum(case when h_brief is null then 0.00 else h_brief end) as h_brief,
                sum(case when h_panne is null then 0.00 else h_panne end) as h_panne,
                sum(case when h_formation is null then 0.00 else h_formation end) as h_formation,
                sum(case when h_relance is null then 0.00 else h_relance end) as h_relance
                 from suivi_production_ta_cdt
                 group by jour,date,day,superviseur_id)as t
                )
        """)
suivi_production_superviseur_cdt()
class suivi_production_employee_cdt(osv.osv):    
    _name = "suivi.production.employee.cdt"
    _description = "Suivi de production par employee"
    _auto = False
    _order ='date desc, employee_id asc'
   
    _columns = {
        'jour':fields.selection(AVAILABLE_JOURS,'Jour',readonly=True),
        'date' : fields.date('Date',readonly=True),
        'day': fields.char('Day',32,),
        'employee_id' : fields.many2one('hr.employee', "Nom et Prénom", readonly=False),
        'cu' : fields.integer('CU'),
        'h_prod_jour':fields.float("HEURES DE PROD JOUR",digit=(6,2)),
        'h_prod_nuit':fields.float("HEURES DE PROD NUIT",digit=(6,2)),
        'h_brief':fields.float("HEURES DE BRIEF",digit=(6,2)),
        'h_panne':fields.float("HEURES DE PANNE",digit=(6,2)),
        'h_relance':fields.float("HEURES DE RELANCE",digit=(6,2)),
        'tx_brief' : fields.char('TAUX DE BRIEF',56),
        'tx_panne' : fields.char('TAUX DE PANNE',56),
        'total_heure':fields.float('TOTAL HEURES'),
        'cadence1':fields.float('CADENCE PANNE INCLUSES'),
        'cadence2':fields.float('CADENCE HORS PANNE'),
        'cadence3':fields.float('CADENCE HORS RELANCE ET PANNES'),  
        'h_formation':fields.float("HEURES DE FORMATIONS",digit=(6,2)),      
       }
    def init(self, cr):
        tools.drop_view_if_exists(cr, 'suivi_production_employee_cdt')
        cr.execute("""
            create or replace view suivi_production_employee_cdt as (
                 select t.id,t.jour,t.date,t.day,t.employee_id,t.cu,t.h_prod_jour,t.h_prod_nuit,t.h_brief,t.h_panne,t.h_relance,t.h_formation,
                case when (t.h_prod_jour+t.h_prod_nuit+t.h_brief+t.h_panne+t.h_relance)=0 then '0.00%' else 
                round(cast(t.h_brief*100.00/(t.h_prod_jour+t.h_prod_nuit+t.h_brief+t.h_panne+t.h_relance) as numeric),2)||'%' end as tx_brief,
                case when (t.h_prod_jour+t.h_prod_nuit+t.h_brief+t.h_panne+t.h_relance)=0 then '0.00%' else 
                round(cast(t.h_panne*100.00/(t.h_prod_jour+t.h_prod_nuit+t.h_brief+t.h_panne+t.h_relance) as numeric),2)||'%' end as tx_panne,
                (t.h_prod_jour+t.h_prod_nuit+t.h_brief+t.h_panne+t.h_relance) as total_heure,
                case when (t.h_prod_jour+t.h_prod_nuit+t.h_brief+t.h_panne+t.h_relance)=0 then 0.00 else 
                round(cast(t.cu*1.00/(t.h_prod_jour+t.h_prod_nuit+t.h_brief+t.h_panne+t.h_relance) as numeric),2) end as cadence1,
                case when (t.h_prod_jour+t.h_prod_nuit+t.h_brief+t.h_relance)=0 then 0.00 else 
                round(cast(t.cu*1.00/(t.h_prod_jour+t.h_prod_nuit+t.h_brief+t.h_relance) as numeric),2) end as cadence2,
                case when (t.h_prod_jour+t.h_prod_nuit+t.h_brief)=0 then 0.00 else 
                round(cast(t.cu*1.00/(t.h_prod_jour+t.h_prod_nuit+t.h_brief) as numeric),2) end as cadence3
                from (select min(id) as id,jour,date,day,employee_id,sum(cu) as cu,
                sum(case when h_prod_jour is null then 0.00 else h_prod_jour end) as h_prod_jour,
                sum(case when h_prod_nuit is null then 0.00 else h_prod_nuit end) as h_prod_nuit,
                sum(case when h_brief is null then 0.00 else h_brief end) as h_brief,
                sum(case when h_panne is null then 0.00 else h_panne end) as h_panne,
                sum(case when h_formation is null then 0.00 else h_formation end) as h_formation,
                sum(case when h_relance is null then 0.00 else h_relance end) as h_relance
                 from suivi_production_ta_cdt
                 group by jour,date,day,employee_id)as t
                )
        """)
suivi_production_employee_cdt()
class production_base_type_cdt(osv.osv):
    _name = 'production.base.type.cdt'
    _order ='name asc'
    
    _columns = {
        'name' : fields.char('TYPE DOSSIER',56,required=True,readonly=False),
        }
    _defaults = {
         }
    _sql_constraints = [
        ('name_uniq', 'unique (name)', 'Le type doit etre unique!')
    ]
production_base_type_cdt()
class production_base_work_cdt(osv.osv):
    _name = 'production.base.work.cdt'
    _order ='date desc'
    
    _columns = {
        'date' : fields.date('Date'),
        'base' : fields.char('Base',size=32),
        }
    _defaults = {'date': lambda *a: time.strftime('%Y-%m-%d'),
         }
    _sql_constraints = [
        ('name_uniq', 'unique (name)', 'Le type doit etre unique!')
    ]
production_base_work_cdt()
class production_base_cdt(osv.osv):
    _name = 'production.base.cdt'
    _order ='is_relance desc'
    
    _columns = {
        'name' : fields.char('Base',56,required=True,readonly=False),
        'type_id'  : fields.many2one('production.base.type.cdt','Type Dossier'),
        'type_tarif' : fields.selection([('invitation','Invitation'),('heure','Heure'),('retour','Retour')],'Type Tarif'),
        'tarif' : fields.float('Tarif'),
        'is_relance' : fields.selection(AVAILABLE_RELANCES, 'Est une base relance?', size=16, readonly=False),
        'cu': fields.many2many('production.status.cdt',
            'production_status_base_cdt_rel',
            'production_base_cdt_id', 'production_status_cdt_id', 'Status CU'),
        'cu_h': fields.float('Objectif CU/H'),
        'production_ids' : fields.one2many('suivi.production.ta.cdt','base_id','Productions',readonly=True)
        }
    _defaults = {
                'type_tarif' : 'invitation' 
         }
    _sql_constraints = [
        ('name_uniq', 'unique (name)', 'Le nom de la base doit etre unique!')
    ]
production_base_cdt()
class production_login_cdt(osv.osv):
    _name = 'production.login.cdt'
    _order ='name asc'
    
    _columns = {
        'name' : fields.char('Login',56,required=True,readonly=False),
        'employee_id': fields.many2one('hr.employee','Nom et Prénom',domain=[('operation_id.name','=','CDT'),('state','in',('open','en_conge'))]),
        }
    _defaults = {
         }
    def write(self, cr, uid,ids, data, context=None):
        obj_id = super(production_login_cdt, self).write(cr, uid,ids, data, context)
        
        cr.execute("select name,employee_id from production_login_cdt where id="+str(ids[0]),(tuple(),))
        for res in cr.fetchall():
            tv=res[0]
            employee_id=res[1]
        if employee_id is not None:
            cr.execute("select superviseur_id from hr_employee where id="+str(employee_id),(tuple(),))
            for res in cr.fetchall():
                cr.execute("update suivi_production_ta_cdt set superviseur_id="+str(res[0])+"\
                            ,employee_id="+str(employee_id)+"\
                            where tv='"+tv+"'",(tuple(),))
                cr.commit()
        else :
            cr.execute("update suivi_production_ta_cdt set superviseur_id=null\
                            ,employee_id=null\
                            where tv='"+tv+"'",(tuple(),))
            cr.commit()
        return obj_id
    _sql_constraints = [
        ('name_uniq', 'unique (name)', 'Le login doit etre unique!')
    ]
production_login_cdt()
class production_retour_cdt(osv.osv):
    _name = 'production.retour.cdt'
    _order = 'year asc,month desc,base desc,login'

    _columns = {
        'year' : fields.integer('Année', required=True),
        'month' : fields.selection(AVAILABLE_MONTH, 'Mois', required=True),
        'base' : fields.char('Base', required=True),
        'login': fields.char('Login', required=True),
        'invitation' : fields.integer('Invitations'),
        'retour' : fields.integer('Retour')
    }
    _defaults = {
    }
production_retour_cdt()
class production_invoice_cdt(osv.osv):
    _name = 'production.invoice.cdt'
    _order ='date asc'
    def _get_name(self,cr,uid,ids, name, args, context=True):
        res = {}
        for obj in self.browse(cr, uid, ids, context=context):
            res[obj.id] = 'Facture ' + datetime.datetime.strptime(obj.date,'%Y-%m-%d').strftime('%d/%m/%Y')
        return res
    def _compute_amount_total(self,cr,uid,ids, name, args, context=True):
        res = {}
        for obj in self.browse(cr, uid, ids, context=context):
            amount_total=0
            for line in obj.invoice_line_ids:
                amount_total += line.price * line.cu_h            
            res[obj.id] = amount_total
        return res
    def _compute_amount_paid_total(self,cr,uid,ids, name, args, context=True):
        res = {}
        for obj in self.browse(cr, uid, ids, context=context):
            amount_paid_total=0
            for line in obj.invoice_line_ids:
                if line.paid== True:
                    amount_paid_total += line.price * line.cu_h            
            res[obj.id] = amount_paid_total
        return res
    def _compute_amount_balance_total(self,cr,uid,ids, name, args, context=True):
        res = {}
        for obj in self.browse(cr, uid, ids, context=context):
            amount_total=amount_paid_total=0
            for line in obj.invoice_line_ids:
                amount_total += line.price * line.cu_h 
                if line.paid== True:
                    amount_paid_total += line.price * line.cu_h            
            res[obj.id] = amount_total-amount_paid_total
        return res
    _columns = {
        'name' : fields.function(fnct=_get_name,string="N° Facture",type='char', store=True),               
        'date' : fields.date('Date', required = True),
        'prod' : fields.integer('Prod' , required = True),
        'retour' : fields.integer('Retour' , required = True),
        'invoice_line_ids' : fields.one2many('production.invoice.line.cdt', 'invoice_id', string="Factures"),        
        'amount_total' : fields.function(_compute_amount_total,string='Montant Total',type='float',),
        'amount_paid_total' : fields.function(_compute_amount_paid_total,string='Montant payé',type='float',),
        'amount_balance_total' : fields.function(_compute_amount_balance_total,string='Différence',type='float')
        }
    _defaults = {
         }
    _sql_constraints = [
        ('name_uniq', 'unique (date)', 'Les factures sont uniques!')
    ]
production_invoice_cdt()
class production_invoice_line_cdt(osv.osv):
    _name = 'production.invoice.line.cdt'
    _order = 'date desc,site,name desc'
    
    def _compute_amount(self,cr,uid,ids, name, args, context=True):
        res = {}
        for obj in self.browse(cr, uid, ids, context=context):
            res[obj.id] = obj.price * obj.cu_h
        return res
    def _get_name(self,cr,uid,ids, name, args, context=True):
        res = {}
        for obj in self.browse(cr, uid, ids, context=context):
            month="01"
            year='001'
            _logger.info(year)
            if obj.year and obj.month:
                name=str(obj.month)+''+datetime.datetime.strptime(obj.invoice_id.date,'%Y-%m-%d').strftime('%m%Y/')+str(obj.year)
            else :
                cr.execute("select lpad((cast(max(right(name,3)) as numeric)+1)::text, 3, '0') from production_invoice_line_cdt\
                            where extract(year from date)=extract(year from date('"+str(obj.invoice_id.date)+"'))\
                            and (right(name,3))~ '^[0-9\.]+$'\
                            and site='"+obj.site+"'\
                            and id<"+str(obj.id))
                for res1 in cr.fetchall():
                    if res1[0]:
                        year=res1[0]
                cr.execute("select lpad((cast(max(left(name,2)) as numeric)+1)::text, 2, '0') from production_invoice_line_cdt\
                            where extract(year from date)=extract(year from date('"+str(obj.invoice_id.date)+"'))\
                            and extract(month from date)=extract(month from date('"+str(obj.invoice_id.date)+"'))\
                            and (left(name,2))~ '^[0-9\.]+$'\
                            and site='"+obj.site+"'\
                            and id<"+str(obj.id))
                for res1 in cr.fetchall():
                    if res1[0]:
                        month=res1[0]
                _logger.info(year)
            
                name=str(month)+''+datetime.datetime.strptime(obj.invoice_id.date,'%Y-%m-%d').strftime('%m%Y/')+str(year)
            res[obj.id] = name
        return res
    def _compute_date(self,cr,uid,ids, name, args, context=True):
        res = {}
        for obj in self.browse(cr, uid, ids, context=context):
            res[obj.id] = obj.invoice_id.date
        return res
    def action_paid_amount(self,cr,uid,ids,context=True):            
        self.write(cr,uid,ids,{'paid': True})
        return True
    def action_impaid_amount(self,cr,uid,ids,context=True):            
        self.write(cr,uid,ids,{'paid': False})
        return True
    def action_print(self, cr, uid, ids, context=None):
        datas = {
                 'model': 'production.invoice.line.cdt',
                 'ids': ids,
                 'form': self.read(cr, uid, ids[0], context=context),
        }
        if self.read(cr, uid, ids[0], context=context)['site']=='CASA':
            return {'type': 'ir.actions.report.xml', 'report_name': 'invoice_cdt_casa', 'datas': datas, 'nodestroy': True}
        else :
            return {'type': 'ir.actions.report.xml', 'report_name': 'invoice_cdt_abidjan', 'datas': datas, 'nodestroy': True}
    def _set_name(self, cr, uid, ids, field_name, field_value, arg, context):
        cr.execute(
            'UPDATE production_invoice_line_cdt '
            'SET name=%s '
            'WHERE id=%s', (field_value, ids)
        )
        return self.write(cr,uid,ids,{'name' : field_value}, context=context)  
    
    _columns = {
        'name' : fields.function(fnct=_get_name,string="N° Facture",type='char', store=True),
        'year' : fields.char('Année',size=3),
        'month' : fields.char('Mois',size=2),
        'date' : fields.function(_compute_date,string="Date",type='date', store=True,required=False),
        'invoice_id' : fields.many2one('production.invoice.cdt','Facture',required=True),
        'site' : fields.selection([('CASA','CASA'),('ABIDJAN','ABIDJAN')],'Site',required=True),
        'dossier_id' : fields.many2one('production.base.cdt','Dossier',required=True),
        'cu' : fields.integer('Nb CU',required=True),
        'cu_h' : fields.float('Nb H/ Nb CU',required=True),
        'price' : fields.float('Prix',required=True),
        'amount' : fields.function(_compute_amount,string="Montant Net",type='float', store=True),
        'paid' : fields.boolean(string="Payé"),
        }
    _defaults = {
         }
production_invoice_line_cdt()
class production_status_cdt(osv.osv):
    _name = 'production.status.cdt'

    _columns = {
        'name' : fields.integer('Status'),
        }
    _defaults = {
         }
    _sql_constraints = [
        ('name_uniq', 'unique (name)', 'Le numero du status doit etre unique!')
    ]
production_base_cdt()
class production_update_cdt(osv.osv):
    _name = 'production.update.cdt'
    
    def update_production(self, cr, uid, ids=True, context=True):
        for obj in self.read(cr, uid, ids, ['date_update'], context=context):
            _logger.info('#########################update_production  :  Debut Mis a jour###################################')
            cr.execute("select to_char(date('"+str(obj['date_update'])+"'),'YYYYMMDD')",(tuple(),))
            for res in cr.fetchall():
                date=res[0]
            _logger.info(date)
            self.pool.get('suivi.production.ta.cdt').get_data_hermes( cr, uid,date, ids, context)
            #self.pool.get('suivi.production.ta.cdt').get_data( cr, uid,date, ids, context)
            self.pool.get('suivi.production.ta.cdt').update_data(cr, uid, date, ids, context)
            #self.pool.get('production.tarification').envoyer_grille_tarifaire( cr, uid, ids, context)
            _logger.info('#########################update_production  :  Fin Mis a jour###################################')
        return True
    _columns = {
        'date_update':fields.date('Date Mis a jour'),
        'heure_update': fields.char('Heure Appel',128),
        'objet': fields.char('Objet',56),
        }
    _defaults = {
                 'date_update': lambda *a: time.strftime('%Y-%m-%d'),
         }
production_update_cdt()
class production_prime_quot_cdt(osv.osv):    
    _name = "production.prime.quot.cdt"
    _description = "Primes Quotidiennes CDT"
    _order = 'date desc'
    
    _columns = {
        'date' : fields.date('Date',required=True),        
        'employee_id': fields.many2one('hr.employee','Nom et Prénom',required=True,domain=[('state','in',('open','en_conge'))]),
        'montant' : fields.float('Montant')       
       }
    _defaults = {
    
    }
production_prime_quot_cdt()
class production_server_cdt(osv.osv):    
    _name = "production.server.cdt"
    _description = "Serveur CDT"
    def test_server_connection(self, cr, uid, ids, context=None):
        for server in self.browse(cr, uid, ids, context=context):
            smtp = False
            try:
                conn = pymssql.connect(host=server.server_host,
                                   user=server.server_user, 
                                   password=server.server_password, 
                                   database=server.data_base)
            except Exception, e:
                raise osv.except_osv(_("Connection test failed!"), _("Here is what we got instead:\n %s") % tools.ustr(e))
            finally:
                conn.close()
        raise osv.except_osv(_("Connection test succeeded!"), _("Everything seems properly set up!"))
    
    
    _columns = {
        'name':fields.selection([('APPEL','APPEL'),('IDENT','IDENT')],'Serveur',required=True),
        'server_host':fields.char('Host',128,required=True),
        'server_user': fields.char('Utilisateur',128,required=True),
        'server_password': fields.char('Mot de passe',128,required=True),
        'data_base': fields.char('BDD Appel',128,required=False),
       }
    _defaults = {
    
    }
production_server_cdt()
class suivi_prime_ca_cdt(osv.osv):    
    _name = "suivi.prime.ca.cdt"
    _description = "Suivi du CA CDT"
    _auto = False
    _order ='date desc'
    _columns = {
        'date' : fields.date('Date'),
        'base_id' : fields.many2one('production.base.cdt', "Base"),                      
        'employee_id' : fields.many2one('hr.employee','Nom et Prénom'),  
        'cu' : fields.integer('CU'),
        'ca' : fields.float('CA'),
        'h_prod':fields.float("H.PROD",digit=(6,2)),
        'h_brief':fields.float("H.BRIEF",digit=(6,2)),
        'h_panne':fields.float("H.PANNE",digit=(6,2)),
        'h_relance':fields.float("H.RELANCE",digit=(6,2)),
       }
    def init(self, cr):
        tools.drop_view_if_exists(cr, 'suivi_prime_ca_cdt')
        cr.execute("""
            create or replace view suivi_prime_ca_cdt as (
                 select row_number() OVER () as id,tab.* from(
                    select date,base_id,employee_id,cu,
                    case when h_prod_jour is null then 0 else h_prod_jour end as h_prod,
                    case when h_brief is null then 0 else h_brief end as h_brief,
                    case when h_panne is null then 0 else h_panne end as h_panne,
                    case when h_relance is null then 0 else h_relance end as h_relance,
                    case when type_tarif='invitation' then cu*tarif 
                         when type_tarif='heure' then (case when h_prod_jour is null then 0 else h_prod_jour end)*tarif end as ca
                    from suivi_production_ta_cdt a
                    inner join production_base_cdt b on a.base_id=b.id
                    left join production_base_type_cdt c on c.id=b.type_id) as tab
                    order by date desc
             )
        """)
suivi_prime_ca_cdt()
class suivi_board_prime_ca_cdt(osv.osv):    
    _name = "suivi.board.prime.ca.cdt"
    _description = "Suivi TBD du CA CDT"
    _order = 'sequence'
    _auto = False

    _columns = {
        'periode' : fields.char('Période',readonly=True),
        'sequence' : fields.integer('Séquence',readonly=True),                    
        'employee_id' : fields.many2one('hr.employee','Nom et Prénom',readonly=True),  
        'cu' : fields.integer('CU',readonly=True),
        'ca' : fields.float('CA',readonly=True),
        'h_prod':fields.float("H.PROD",digit=(6,2),readonly=True),
        'h_panne':fields.float("H.PANNE",digit=(6,2),readonly=True),
        'h_relance':fields.float("H.RELANCE",digit=(6,2),readonly=True),
        'ca_h_1':fields.float("CA / H HORS RELANCE ET PANNES",digit=(6,2),readonly=True),
        'ca_h_2':fields.float("CA / H RELANCE INCLUSE",digit=(6,2),readonly=True),
        'ca_h_3':fields.float("CA / H RELANCE ET PANNE INCLUSE",digit=(6,2),readonly=True),
       }
    def init(self, cr):
        tools.drop_view_if_exists(cr, 'suivi_board_prime_ca_cdt')
        cr.execute("""
            create or replace view suivi_board_prime_ca_cdt as (
                 select row_number() OVER () as id,tab.* from(
                select 1 as sequence,'La veille' as periode,employee_id,sum(cu) as cu,sum(ca) as ca,sum(h_prod) as h_prod,sum(h_panne) as h_panne,sum(h_relance) as h_relance,
                round(cast(case when sum(h_prod)=0 then 0 else sum(ca)/sum(h_prod) end as numeric),2) as ca_h_1,
                round(cast(case when sum(h_prod+h_relance)=0 then 0 else sum(ca)/(sum(h_prod+h_relance)) end as numeric),2) as ca_h_2,
                round(cast(case when sum(h_prod+h_panne+h_relance)=0 then 0 else sum(ca)/sum(h_prod+h_panne+h_relance) end as numeric),2) as ca_h_3
                 from suivi_prime_ca_cdt
                where date=date(now())-1
                group by employee_id
                union
                select 2 as sequence,'Semaine en cours' as type,employee_id,sum(cu),sum(ca) as ca,sum(h_prod) as h_prod,sum(h_panne) as h_panne,sum(h_relance) as h_relance,
                round(cast(case when sum(h_prod)=0 then 0 else sum(ca)/sum(h_prod) end as numeric),2) as ca_h_1,
                round(cast(case when sum(h_prod+h_relance)=0 then 0 else sum(ca)/(sum(h_prod+h_relance)) end as numeric),2) as ca_h_2,
                round(cast(case when sum(h_prod+h_panne+h_relance)=0 then 0 else sum(ca)/sum(h_prod+h_panne+h_relance) end as numeric),2) as ca_h_3
                 from suivi_prime_ca_cdt
                where extract(year from date)=extract(year from now())
                and extract(month from date)=extract(month from now())
                and extract(week from date)=extract(week from now())
                group by employee_id
                union
                select 3 as sequence,'Mois en cours' as type,employee_id,sum(cu),sum(ca) as ca,sum(h_prod) as h_prod,sum(h_panne) as h_panne,sum(h_relance) as h_relance,
                round(cast(case when sum(h_prod)=0 then 0 else sum(ca)/sum(h_prod) end as numeric),2) as ca_h_1,
                round(cast(case when sum(h_prod+h_relance)=0 then 0 else sum(ca)/(sum(h_prod+h_relance)) end as numeric),2) as ca_h_2,
                round(cast(case when sum(h_prod+h_panne+h_relance)=0 then 0 else sum(ca)/sum(h_prod+h_panne+h_relance) end as numeric),2) as ca_h_3
                 from suivi_prime_ca_cdt
                where extract(year from date)=extract(year from now())
                and extract(month from date)=extract(month from now())
                group by employee_id) as tab order by sequence

             )
        """)
suivi_board_prime_ca_cdt()
class suivi_board_stats_cdt(osv.osv):    
    _name = "suivi.board.stats.cdt"
    _description = "Suivi TBD du Stats CDT"
    _order = 'sequence'
    _auto = False

    _columns = {
        
        'sequence' : fields.integer('Séquence',readonly=True),                    
        'employee_id' : fields.many2one('hr.employee','Nom et Prénom',readonly=True),  
        'information' : fields.char('Information',readonly=True),
        'data' : fields.float('Donnée',readonly=True),
         }
    def init(self, cr):
        tools.drop_view_if_exists(cr, 'suivi_board_stats_cdt')
        cr.execute("""
            create or replace view suivi_board_stats_cdt as (
                  select row_number() OVER () as id,tab.* from(
         select 1 as sequence,employee_id,'CA' as information,sum(ca) as data from suivi_prime_ca_cdt
        where extract(year from date)=extract(year from now())
        and extract(month from date)=extract(month from now())
        group by employee_id
        union
         select 2 as sequence,employee_id,'HEURES PROD' as information,sum(h_prod) as data from suivi_prime_ca_cdt
        where extract(year from date)=extract(year from now())
        and extract(month from date)=extract(month from now())
        group by employee_id
        union
         select 3 as sequence,employee_id,'HEURES RELANCE NON FACTUREE' as information,sum(h_relance) as data from suivi_prime_ca_cdt
        where extract(year from date)=extract(year from now())
        and extract(month from date)=extract(month from now())
        group by employee_id
        union
         select 4 as sequence,employee_id,'HEURES BRIEF' as information,sum(h_brief) as data from suivi_prime_ca_cdt
        where extract(year from date)=extract(year from now())
        and extract(month from date)=extract(month from now())
        group by employee_id
        union
         select 5 as sequence,employee_id,'HEURES PANNE' as information,sum(h_panne) as data from suivi_prime_ca_cdt
        where extract(year from date)=extract(year from now())
        and extract(month from date)=extract(month from now())
        group by employee_id
        ) as tab
             )
        """)
suivi_board_stats_cdt()
class suivi_board_retour_cdt(osv.osv):
    _name = "suivi.board.retour.cdt"
    _description = "Suivi TBD du retours CDT"
    _order = 'nbr,taux_retour desc'
    _auto = False

    _columns = {
        'nbr' : fields.integer('Nb'),
        'employee_id': fields.many2one('hr.employee', 'Nom et Prénom', readonly=True),
        'superviseur_id': fields.many2one('hr.employee', 'Equipe', readonly=True),
        'year' : fields.integer('Année', required=True),
        'month' : fields.selection(AVAILABLE_MONTH, 'Mois', required=True),
        'invitation' : fields.integer('Invitations'),
        'retour' : fields.integer('Retour'),
        'taux_retour' : fields.float('Tx de retour'),
        'taux_retour_text' : fields.char('Tx de retour')
    }

    def init(self, cr):
        tools.drop_view_if_exists(cr, 'suivi_board_retour_cdt')
        cr.execute("""
            create or replace view suivi_board_retour_cdt as (
                select row_number() OVER () as id,tab.* from(select 1 as nbr,year,month,emp.superviseur_id,b.employee_id,sum(invitation) as invitation,sum(retour) as retour,
                    round(cast(case when sum(invitation)=0 then 0 else sum(retour)*100.00/sum(invitation) end as numeric),2) as taux_retour,
                    round(cast(case when sum(invitation)=0 then 0 else sum(retour)*100.00/sum(invitation) end as numeric),2)||'%' as taux_retour_text
                    from production_retour_cdt a
                    inner join production_login_cdt b on a.login=b.name
                    inner join hr_employee emp on b.employee_id=emp.id
                    where year=extract(year from now())
                    and month=extract(month from now())
                    group by year,month,emp.superviseur_id,employee_id
                union
                select 2 as nbr,year,month,emp.superviseur_id,emp.superviseur_id,sum(invitation) as invitation,sum(retour) as retour,
                    round(cast(case when sum(invitation)=0 then 0 else sum(retour)*100.00/sum(invitation) end as numeric),2) as taux_retour,
                    round(cast(case when sum(invitation)=0 then 0 else sum(retour)*100.00/sum(invitation) end as numeric),2)||'%' as taux_retour_text
                    from production_retour_cdt a
                    inner join production_login_cdt b on a.login=b.name
                    inner join hr_employee emp on b.employee_id=emp.id
                    where year=extract(year from now())
                    and month=extract(month from now())
                    group by year,month,emp.superviseur_id
                union
                select 3 as nbr,year,month,(select id from hr_employee where complete_name like '%MORSAD%' and state='open') as employee_id,
                    (select id from hr_employee where complete_name like '%MORSAD%' and state='open')as superviseur_id,
                    sum(invitation) as invitation,sum(retour) as retour,
                    round(cast(case when sum(invitation)=0 then 0 else sum(retour)*100.00/sum(invitation) end as numeric),2) as taux_retour,
                    round(cast(case when sum(invitation)=0 then 0 else sum(retour)*100.00/sum(invitation) end as numeric),2)||'%' as taux_retour_text
                    from production_retour_cdt a
                    inner join production_login_cdt b on a.login=b.name
                    inner join hr_employee emp on b.employee_id=emp.id
                    where year=extract(year from now())
                    and month=extract(month from now())
                    group by year,month
                    ) as tab
             )
        """)
suivi_board_retour_cdt()
class suivi_board_sup_ca_cdt(osv.osv):    
    _name = "suivi.board.sup.ca.cdt"
    _description = "Suivi TBD du SUP CDT"
    _order = 'type,superviseur_id,type2,ca_h_m desc'
    _auto = False

    _columns = {
        'type' : fields.char('Type',readonly=True),
        'type2' : fields.integer('Type',readonly=True),
        'employee_id' : fields.many2one('hr.employee','Nom et Prénom',readonly=True),  
        'superviseur_id' : fields.many2one('hr.employee','Superviseur',readonly=True),  
        'ca_h_j':fields.float("CA / H Veille",digit=(6,2),readonly=True),
        'ca_h_s':fields.float("CA / H Semaine en cours",digit=(6,2),readonly=True),
        'ca_h_m':fields.float("CA / H Mois en cours",digit=(6,2),readonly=True),
       }
    def init(self, cr):
        tools.drop_view_if_exists(cr, 'suivi_board_sup_ca_cdt')
        cr.execute("""
            create or replace view suivi_board_sup_ca_cdt as (
                 select row_number() OVER () as id,tab.* from(
select 'ss_total' as type,1 as type2,employee_id,emp.superviseur_id,
                round(cast(case when sum(case when date=date(now())-1 then h_prod else 0 end)=0 then 0 
                else sum(case when date=date(now())-1 then ca else 0 end)/sum(case when date=date(now())-1 then h_prod else 0 end) end as numeric),2) as ca_h_j,
        round(cast(case when sum(case when extract(year from date)=extract(year from now()) and extract(week from date)=extract(week from now()) then h_prod else 0 end)=0 then 0 
                else sum(case when extract(year from date)=extract(year from now()) and extract(week from date)=extract(week from now()) then ca else 0 end)/sum(case when extract(year from date)=extract(year from now()) 
                and extract(week from date)=extract(week from now()) then h_prod else 0 end) end as numeric),2) as ca_h_s,
        round(cast(case when sum(case when extract(year from date)=extract(year from now()) and extract(month from date)=extract(month from now()) then h_prod else 0 end)=0 then 0 
                else sum(case when extract(year from date)=extract(year from now()) and extract(month from date)=extract(month from now()) then ca else 0 end)/sum(case when extract(year from date)=extract(year from now()) 
                and extract(month from date)=extract(month from now()) then h_prod else 0 end) end as numeric),2) as ca_h_m        
                 from suivi_prime_ca_cdt a inner join hr_employee emp on a.employee_id=emp.id
                 where extract(year from date)=extract(year from now()) and extract(month from date)=extract(month from now())
                group by employee_id,emp.superviseur_id
union

                select 'ss_total' as type,2 as type2,emp.superviseur_id as employee_id,emp.superviseur_id as superviseur_id,
                round(cast(case when sum(case when date=date(now())-1 then h_prod else 0 end)=0 then 0 
                else sum(case when date=date(now())-1 then ca else 0 end)/sum(case when date=date(now())-1 then h_prod else 0 end) end as numeric),2) as ca_h_j,
        round(cast(case when sum(case when extract(year from date)=extract(year from now()) and extract(week from date)=extract(week from now()) then h_prod else 0 end)=0 then 0 
                else sum(case when extract(year from date)=extract(year from now()) and extract(week from date)=extract(week from now()) then ca else 0 end)/sum(case when extract(year from date)=extract(year from now()) 
                and extract(week from date)=extract(week from now()) then h_prod else 0 end) end as numeric),2) as ca_h_s,
        round(cast(case when sum(case when extract(year from date)=extract(year from now()) and extract(month from date)=extract(month from now()) then h_prod else 0 end)=0 then 0 
                else sum(case when extract(year from date)=extract(year from now()) and extract(month from date)=extract(month from now()) then ca else 0 end)/sum(case when extract(year from date)=extract(year from now()) 
                and extract(month from date)=extract(month from now()) then h_prod else 0 end) end as numeric),2) as ca_h_m        
                 from suivi_prime_ca_cdt a inner join hr_employee emp on a.employee_id=emp.id
                 where extract(year from date)=extract(year from now()) and extract(month from date)=extract(month from now())
                group by emp.superviseur_id
 union

                select 'vtotal' as type,3 as type2,(select id from hr_employee where complete_name like '%MORSAD%' and state='open') as employee_id,
            (select id from hr_employee where complete_name like '%MORSAD%' and state='open')as superviseur_id,
                round(cast(case when sum(case when date=date(now())-1 then h_prod else 0 end)=0 then 0 
                else sum(case when date=date(now())-1 then ca else 0 end)/sum(case when date=date(now())-1 then h_prod else 0 end) end as numeric),2) as ca_h_j,
        round(cast(case when sum(case when extract(year from date)=extract(year from now()) and extract(week from date)=extract(week from now()) then h_prod else 0 end)=0 then 0 
                else sum(case when extract(year from date)=extract(year from now()) and extract(week from date)=extract(week from now()) then ca else 0 end)/sum(case when extract(year from date)=extract(year from now()) 
                and extract(week from date)=extract(week from now()) then h_prod else 0 end) end as numeric),2) as ca_h_s,
        round(cast(case when sum(case when extract(year from date)=extract(year from now()) and extract(month from date)=extract(month from now()) then h_prod else 0 end)=0 then 0 
                else sum(case when extract(year from date)=extract(year from now()) and extract(month from date)=extract(month from now()) then ca else 0 end)/sum(case when extract(year from date)=extract(year from now()) 
                and extract(month from date)=extract(month from now()) then h_prod else 0 end) end as numeric),2) as ca_h_m        
                 from suivi_prime_ca_cdt a inner join hr_employee emp on a.employee_id=emp.id
                 where extract(year from date)=extract(year from now()) and extract(month from date)=extract(month from now())
           ) as tab



             )
        """)
suivi_board_sup_ca_cdt()
class reporting(osv.osv):
    _name = 'reporting'
    _inherit = 'reporting'
    _columns = {'type_tarif' : fields.selection([('invitation','Invitation'),('heure','Heure'),('retour','Retour'),('invitation_heure','Invitation+Heure'),('total', 'Total')],'Type Tarif'),
                'type_cdt' : fields.selection([('TOTAL', 'TOTAL'), ('EQUIPE', 'EQUIPE')],'Type'),
                'superviseur_id' : fields.many2one('hr.employee','Equipe',domain=[('category_id.name','in',('SUP','Chef de plateau')),('state','in',('open','en_conge')),('operation_id.name','=','CDT')]),
        
        }
    def generer_stats_cdt(self, cr, uid,ids=True, context=None):
        reload(sys)
        sys.setdefaultencoding("UTF8")       
        for obj in self.read(cr, uid, ids, ['annee','mois','type_cdt', 'type_tarif' , 'superviseur_id'], context=context):
            mois=obj['mois']
            annee=obj['annee']
            type_cdt = obj['type_cdt']
            type_tarif = obj['type_tarif']
            superviseur_id = None

            if obj['superviseur_id']:
                superviseur_id = obj['superviseur_id'][0]
        fichier=self.generer_stats_cdt_(cr, uid,type_cdt,type_tarif,superviseur_id,mois,annee, ids, context)
        return self.get_return(cr, uid,fichier, ids, context)
    def generer_stats_cdt_(self, cr, uid,type_cdt,type_tarif,superviseur_id,mois,annee,ids=True, context=None):
        cr.execute("update suivi_production_ta_cdt set tv=replace(tv,'''','') where tv like '%''%'")
        cr.commit()
        reload(sys)
        sys.setdefaultencoding("UTF8")  
        le_mois=''
        le_mois=self.get_le_mois(mois)
        date_deb=str(annee)+'-'+str(le_mois)+'-01'

        requete = " and type_tarif= '"+str(type_tarif)+ "'  "
        if type_tarif == 'total':
            requete = ''
        if type_tarif == 'invitation_heure':
            requete = " and type_tarif in ('invitation', 'heure')  "
        filename ='TOTAL'
        if type_cdt == 'EQUIPE':
            requete+= ' and a.superviseur_id = ' +str(superviseur_id)+' '
            cr.execute('select complete_name from hr_employee where id={0}'.format(superviseur_id))
            for res in cr.fetchall():
                complete_name=res[0]
            filename ='EQUIPE '+complete_name+'_'
        ###################### Export Excel Trame CE###################
        fichier="STATS CDT_"+filename+mois+" "+str(annee)+"_"+str(time.strftime('%H%M%S'))+".xlsx"
        workbook = xlsxwriter.Workbook(sortie + fichier)
        style= workbook.add_format({   'text_wrap' : True,'bold' :1,  'align' : 'center', 'valign' : 'vcenter','font_size' : 6, 'border' : 1,'font_name' : 'Arial'})
        style_number= workbook.add_format({ 'num_format' : '#,##0.00',  'text_wrap' : True,'bold' :1,  'align' : 'center', 'valign' : 'vcenter','font_size' : 6,'border' : 1,'font_name' : 'Arial' })
        style_pourcentage= workbook.add_format({ 'num_format' : '0.00%',  'text_wrap' : True,'bold' :1, 'align' : 'center', 'valign' : 'vcenter','font_size' : 6,'border' : 1,'font_name' : 'Arial'})
        style_yellow= workbook.add_format({   'text_wrap' : True,'bold' :1, 'align' : 'center', 'valign' : 'vcenter','font_size' : 6,  'bg_color':  '#FFFFCC' ,   'border' : 1,'font_name' : 'Arial'}) 
        style_orange= workbook.add_format({   'text_wrap' : True,'bold' :1, 'align' : 'center', 'valign' : 'vcenter','font_size' : 6,  'bg_color':  'orange' ,   'border' : 1,'font_name' : 'Arial' }) 
        style_blue= workbook.add_format({   'text_wrap' : True,'bold' :1, 'align' : 'center', 'valign' : 'vcenter','font_size' : 6,  'bg_color':  '#DBE5F1' ,   'border' : 1,'font_name' : 'Arial' })
        cr.execute("""select t.date_part,date_part as classement from (select
                        distinct extract(week from date(date_trunc('month', date('{0}')))+generate_series
                        (0,(select cast(date_part('day',date_now + '1 month - 1 day'::interval)-1 AS integer)
                        FROM ( select date(date_trunc('month', date('{0}'))) as date_now) months 
                        ))))as t order by classement asc""".format(date_deb))
        for res_semaine in cr.fetchall():
            semaine=int(res_semaine[0])
            cr.execute("""select test1,date(test2)  as test2,to_char(date(test2),'DD') from
                            (select 
                            extract(week from date(date_trunc('month', date('{0}')))+generate_series
                            (0,(select cast(date_part('day',date_now + '1 month - 1 day'::interval)-1 AS integer)
                            FROM ( select date(date_trunc('month', date('{0}'))) as date_now) months))) as test1,
                            to_char(date(date_trunc('month', date('{0}')))+generate_series
                            (0,(select cast(date_part('day',date_now + '1 month - 1 day'::interval)-1 AS integer)
                            FROM ( select date(date_trunc('month', date('{0}'))) as date_now) months 
                            )), 'YYYY-MM-DD')  as test2) as table1
                            where test1='{1}'""".format(date_deb,semaine))
            for res_date in cr.fetchall():
                date=res_date[1] 
                feuille = workbook.add_worksheet(res_date[2])
                requete_date = " where a.date='"+date+"' "+requete
                cr.execute("select count(*) from suivi_production_ta_cdt a inner join production_base_cdt b on a.base_id=b.id {0}".format(requete_date))
                for res_test in cr.fetchall():
                    
                    if res_test[0]>0:
                        feuille.set_row(0, 40)
                        compt=5
                        while compt<100:
                            feuille.set_row(compt, 11.25)
                            compt+=1
                        feuille.set_column('A:A', 2)
                        feuille.set_column('B:C', 30)
                        feuille.write('B4','NOM',style)    
                        feuille.write('C4','PRENOM',style) 
                        feuille.write('C1','NOM DOSSIER',style)
                        feuille.write('C2','TYPE DOSSIER',style)
                        feuille.write('C3','TARIF',style)
                        x=3
                        y=1
                        cr.execute("""select distinct a.base_id,b.name,c.name,b.tarif,b.type_tarif from suivi_production_ta_cdt a
                                        inner join production_base_cdt b on a.base_id=b.id
                                        left join production_base_type_cdt c on c.id=b.type_id
                                        {0} order by b.name""".format(requete_date))
                        for res in cr.fetchall():
                            feuille.set_column(self.get_nom_colonne(x)+':'+self.get_nom_colonne(x+1), 4.71)
                            
                            feuille.merge_range(self.get_nom_colonne(x)+str(y)+':'+self.get_nom_colonne(x+1)+str(y),res[1],style)
                            feuille.merge_range(self.get_nom_colonne(x)+str(y+1)+':'+self.get_nom_colonne(x+1)+str(y+1),res[2],style)
                            feuille.merge_range(self.get_nom_colonne(x)+str(y+2)+':'+self.get_nom_colonne(x+1)+str(y+2),res[3],style)
                            
                            feuille.write(self.get_nom_colonne(x)+str(y+3),'H PRD',style_yellow)
                            feuille.write(self.get_nom_colonne(x+1)+str(y+3),'CU',style_blue)
                            x=x+2
                        feuille.set_column(self.get_nom_colonne(x)+':'+self.get_nom_colonne(x), 10.29)
                        feuille.set_column(self.get_nom_colonne(x+1)+':'+self.get_nom_colonne(x+1), 5)
                        feuille.set_column(self.get_nom_colonne(x+2)+':'+self.get_nom_colonne(x+2), 8)
                        feuille.set_column(self.get_nom_colonne(x+3)+':'+self.get_nom_colonne(x+15), 16)
                            
                        feuille.write(self.get_nom_colonne(x)+str(y+3),'HEURES PROD',style_yellow)
                        feuille.write(self.get_nom_colonne(x+1)+str(y+3),'CA',style)
                        feuille.write(self.get_nom_colonne(x+2)+str(y+3),'CUMUL CU',style_blue)
                        feuille.write(self.get_nom_colonne(x+3)+str(y+3),'CA / H HORS RELANCE ET PANNES',style)
                        feuille.write(self.get_nom_colonne(x+4)+str(y+3),'CADENCE HORS RELANCE ET PANNES',style)
                        feuille.write(self.get_nom_colonne(x+5)+str(y+3),'HEURES PANNE',style)
                        feuille.write(self.get_nom_colonne(x+6)+str(y+3),'HEURES RELANCE',style)
                        feuille.write(self.get_nom_colonne(x+7)+str(y+3),'TOTAL HEURES REMUNEREES',style)
                        feuille.write(self.get_nom_colonne(x+8)+str(y+3),'TAUX PANNE',style)
                        feuille.write(self.get_nom_colonne(x+9)+str(y+3),'CA / H RELANCE INCLUSE',style)
                        feuille.write(self.get_nom_colonne(x+10)+str(y+3),'CADENCE RELANCE INCLUSE',style)
                        feuille.write(self.get_nom_colonne(x+11)+str(y+3),'CA / H RELANCE ET PANNE INCLUSE',style)
                        feuille.write(self.get_nom_colonne(x+12)+str(y+3),'CADENCE RELANCE ET PANNE INCLUSE',style)
                        feuille.write(self.get_nom_colonne(x+14)+str(y+3),'H NON REMUNEREES (20 h FORMATION)',style)
                        feuille.write(self.get_nom_colonne(x+15)+str(y+3),'TOTAL HEURE',style)
                        ##################### Remplir Nom et Prénom
                        y_deb=y+4
                        y=y_deb
                        x=0
                        cr.execute("""select row_number() OVER (),name_related,prenom from 
                                (select distinct name_related,prenom from suivi_production_ta_cdt a
                                inner join production_base_cdt c on a.base_id=c.id
                                left join hr_employee b on a.employee_id=b.id 
                                {0} and employee_id is not null 
                                order by name_related,prenom) as t""".format(requete_date))
                    
                        for res in cr.fetchall():
                            feuille.write(self.get_nom_colonne(x)+str(y),res[0],style) 
                            feuille.write(self.get_nom_colonne(x+1)+str(y),res[1],style) 
                            feuille.write(self.get_nom_colonne(x+2)+str(y),res[2],style) 
                            y=y+1
                        feuille.write(self.get_nom_colonne(x+2)+str(y+1),'CA DOSSIER',style) 
                        feuille.write(self.get_nom_colonne(x+2)+str(y+2),'CA / H DOSSIER',style) 
                        feuille.write(self.get_nom_colonne(x+2)+str(y+3),'CU / H DOSSIER',style) 
                        ########################## Remplir Donnée
                        x=3
                        y=y_deb
                        cr.execute("""select distinct employee_id,name_related,prenom from suivi_production_ta_cdt a 
                                        inner join production_base_cdt c on a.base_id=c.id
                                        left join hr_employee b on a.employee_id=b.id
                                        {0} and employee_id is not null order by name_related,prenom""".format(requete_date))
                        for res in cr.fetchall():
                            employee_id=res[0]
                            h_prod_text=ca_text=cu_text=''
                            cr.execute("""select a.base_id,b.name,b.type_tarif
                                        from suivi_production_ta_cdt a
                                        inner join production_base_cdt b on a.base_id=b.id
                                        {0} group by a.base_id,b.name,b.type_tarif order by b.name asc""".format(requete_date))
                            for res2 in cr.fetchall():
                                base_id = res2[0]
                                type_tarif=res2[2]
                                h_prod=cu=''
                                cr.execute("""select sum(case when h_prod_jour is null then 0 else h_prod_jour end) as h_prod,
                                            sum(cu) as invitation
                                            from suivi_production_ta_cdt a
                                            inner join production_base_cdt c on a.base_id=c.id
                                            {0} and employee_id = {1} and base_id={2}""".format(requete_date, employee_id, base_id))
                                for res1 in cr.fetchall():
                                    h_prod=res1[0]
                                    cu=res1[1]
                                h_prod_text+='+'+self.get_nom_colonne(x)+str(y)
                                if type_tarif=='invitation':
                                    ca_text+='+($'+self.get_nom_colonne(x)+'$3*'+self.get_nom_colonne(x+1)+str(y)+')'
                                if type_tarif=='heure':
                                    ca_text+='+($'+self.get_nom_colonne(x)+'$3*'+self.get_nom_colonne(x)+str(y)+')'
                                cu_text+='+'+self.get_nom_colonne(x+1)+str(y)
                                feuille.write(self.get_nom_colonne(x)+str(y),h_prod,style_yellow)
                                feuille.write(self.get_nom_colonne(x+1)+str(y),cu,style_blue)
                                x=x+2
                            cr.execute("""select sum(case when h_panne is null then 0 else h_panne end) as h_panne,
                                        sum(case when h_formation is null then 0 else h_formation end) as h_formation,
                                        sum(case when h_relance is null then 0 else h_relance end) as h_relance
                                        from suivi_production_ta_cdt a
                                        inner join production_base_cdt b on a.base_id=b.id
                                        {0} and employee_id = {1} """.format(requete_date, employee_id))
                            for res1 in cr.fetchall():
                                h_panne=res1[0]
                                h_relance=res1[2]
                                h_formation=res1[1]
                            feuille.write_formula(self.get_nom_colonne(x)+str(y),h_prod_text[1:],style_yellow)
                            feuille.write_formula(self.get_nom_colonne(x+1)+str(y),ca_text[1:],style)
                            feuille.write_formula(self.get_nom_colonne(x+2)+str(y),cu_text[1:],style_blue)
                            feuille.write_formula(self.get_nom_colonne(x+3)+str(y),self.get_nom_colonne(x+1)+str(y)+'/'+self.get_nom_colonne(x)+str(y),style_number)
                            feuille.write_formula(self.get_nom_colonne(x+4)+str(y),self.get_nom_colonne(x+2)+str(y)+'/'+self.get_nom_colonne(x)+str(y),style_number)
                            feuille.write(self.get_nom_colonne(x+5)+str(y),h_panne,style)
                            feuille.write(self.get_nom_colonne(x+6)+str(y),h_relance,style)
                            feuille.write_formula(self.get_nom_colonne(x+7)+str(y),self.get_nom_colonne(x)+str(y)+'+'+self.get_nom_colonne(x+5)+str(y)+'+'+self.get_nom_colonne(x+6)+str(y),style_number)
                            feuille.write_formula(self.get_nom_colonne(x+8)+str(y),self.get_nom_colonne(x+5)+str(y)+'/'+self.get_nom_colonne(x+7)+str(y),style_pourcentage)
                            feuille.write_formula(self.get_nom_colonne(x+9)+str(y),self.get_nom_colonne(x+1)+str(y)+'/('+self.get_nom_colonne(x)+str(y)+'+'+self.get_nom_colonne(x+6)+str(y)+')',style_number)
                            feuille.write_formula(self.get_nom_colonne(x+10)+str(y),self.get_nom_colonne(x+2)+str(y)+'/('+self.get_nom_colonne(x)+str(y)+'+'+self.get_nom_colonne(x+6)+str(y)+')',style_number)
                            feuille.write_formula(self.get_nom_colonne(x+11)+str(y),self.get_nom_colonne(x+1)+str(y)+'/'+self.get_nom_colonne(x+7)+str(y),style_number)
                            feuille.write_formula(self.get_nom_colonne(x+12)+str(y),self.get_nom_colonne(x+2)+str(y)+'/'+self.get_nom_colonne(x+7)+str(y),style_number)
                            feuille.write(self.get_nom_colonne(x+14)+str(y),h_formation,style)
                            feuille.write_formula(self.get_nom_colonne(x+15)+str(y),self.get_nom_colonne(x+14)+str(y)+'+'+self.get_nom_colonne(x+7)+str(y),style_number)
                        
                            y=y+1
                            x=3
                        nb=0
                        cr.execute("""select a.base_id,(select count(distinct employee_id) from suivi_production_ta_cdt a  inner join production_base_cdt c on a.base_id=c.id
                                        {0} and employee_id is not null),b.type_tarif,b.name
                                        from suivi_production_ta_cdt a inner join production_base_cdt b on a.base_id=b.id
                                        {0} group by a.base_id,b.type_tarif,b.name order by b.name""".format(requete_date))
                        for res in cr.fetchall():
                            nb=res[1]
                            type_tarif=res[2]
                            feuille.write_formula(self.get_nom_colonne(x)+str(y),'SUM('+self.get_nom_colonne(x)+str(y_deb)+':'+self.get_nom_colonne(x)+str(y_deb+res[1]-1)+')',style_yellow)
                            feuille.write_formula(self.get_nom_colonne(x+1)+str(y),'SUM('+self.get_nom_colonne(x+1)+str(y_deb)+':'+self.get_nom_colonne(x+1)+str(y_deb+res[1]-1)+')',style_blue)
                            
                            feuille.merge_range(self.get_nom_colonne(x)+str(y+1)+':'+self.get_nom_colonne(x+1)+str(y+1),'',style)
                            feuille.merge_range(self.get_nom_colonne(x)+str(y+2)+':'+self.get_nom_colonne(x+1)+str(y+2),'',style)
                            feuille.merge_range(self.get_nom_colonne(x)+str(y+3)+':'+self.get_nom_colonne(x+1)+str(y+3),'',style)
                            if type_tarif=='invitation' :
                                feuille.write_formula(self.get_nom_colonne(x)+str(y+1),self.get_nom_colonne(x)+'3*'+str(self.get_nom_colonne(x+1)+str(y)),style_number)
                            if type_tarif=='heure' :
                                feuille.write_formula(self.get_nom_colonne(x)+str(y+1),self.get_nom_colonne(x)+'3*'+str(self.get_nom_colonne(x)+str(y)),style_number)
                            if type_tarif=='retour' :
                                feuille.write_formula(self.get_nom_colonne(x)+str(y+1),self.get_nom_colonne(x)+'3*0',style_number)
                            feuille.write_formula(self.get_nom_colonne(x)+str(y+2),self.get_nom_colonne(x)+str(y+1)+'/'+self.get_nom_colonne(x)+str(y),style_number)
                            feuille.write_formula(self.get_nom_colonne(x)+str(y+3),self.get_nom_colonne(x+1)+str(y)+'/'+self.get_nom_colonne(x)+str(y),style_number)
                            x=x+2
                        feuille.write_formula(self.get_nom_colonne(x)+str(y),'SUM('+self.get_nom_colonne(x)+str(y_deb)+':'+self.get_nom_colonne(x)+str(y_deb+nb-1)+')',style_yellow)
                        feuille.write_formula(self.get_nom_colonne(x+1)+str(y),'SUM('+self.get_nom_colonne(x+1)+str(y_deb)+':'+self.get_nom_colonne(x+1)+str(y_deb+nb-1)+')',style)
                        feuille.write_formula(self.get_nom_colonne(x+2)+str(y),'SUM('+self.get_nom_colonne(x+2)+str(y_deb)+':'+self.get_nom_colonne(x+2)+str(y_deb+nb-1)+')',style_blue)
                        feuille.write_formula(self.get_nom_colonne(x+3)+str(y),self.get_nom_colonne(x+1)+str(y)+'/'+self.get_nom_colonne(x)+str(y),style_number)
                        feuille.write_formula(self.get_nom_colonne(x+4)+str(y),self.get_nom_colonne(x+2)+str(y)+'/'+self.get_nom_colonne(x)+str(y),style_number)
                        feuille.write_formula(self.get_nom_colonne(x+5)+str(y),'SUM('+self.get_nom_colonne(x+5)+str(y_deb)+':'+self.get_nom_colonne(x+5)+str(y_deb+nb-1)+')',style)
                        feuille.write_formula(self.get_nom_colonne(x+6)+str(y),'SUM('+self.get_nom_colonne(x+6)+str(y_deb)+':'+self.get_nom_colonne(x+6)+str(y_deb+nb-1)+')',style)
                        feuille.write_formula(self.get_nom_colonne(x+7)+str(y),self.get_nom_colonne(x)+str(y)+'+'+self.get_nom_colonne(x+5)+str(y)+'+'+self.get_nom_colonne(x+6)+str(y),style_number)
                        feuille.write_formula(self.get_nom_colonne(x+8)+str(y),self.get_nom_colonne(x+5)+str(y)+'/'+self.get_nom_colonne(x+7)+str(y),style_pourcentage)
                        feuille.write_formula(self.get_nom_colonne(x+9)+str(y),self.get_nom_colonne(x+1)+str(y)+'/('+self.get_nom_colonne(x)+str(y)+'+'+self.get_nom_colonne(x+6)+str(y)+')',style_number)
                        feuille.write_formula(self.get_nom_colonne(x+10)+str(y),self.get_nom_colonne(x+2)+str(y)+'/('+self.get_nom_colonne(x)+str(y)+'+'+self.get_nom_colonne(x+6)+str(y)+')',style_number)
                        feuille.write_formula(self.get_nom_colonne(x+11)+str(y),self.get_nom_colonne(x+1)+str(y)+'/'+self.get_nom_colonne(x+7)+str(y),style_number)
                        feuille.write_formula(self.get_nom_colonne(x+12)+str(y),self.get_nom_colonne(x+2)+str(y)+'/'+self.get_nom_colonne(x+7)+str(y),style_number)
                        feuille.write_formula(self.get_nom_colonne(x+14)+str(y),'SUM('+self.get_nom_colonne(x+14)+str(y_deb)+':'+self.get_nom_colonne(x+14)+str(y_deb+nb-1)+')',style)
                        feuille.write_formula(self.get_nom_colonne(x+15)+str(y),self.get_nom_colonne(x+14)+str(y)+'+'+self.get_nom_colonne(x+7)+str(y),style_number)
                        
            feuille = workbook.add_worksheet('S'+str(semaine)) 
            requete_week = " where extract(year from a.date)="+str(annee)+" and extract(week from a.date)="+str(semaine)+" \
                        and extract(month from a.date)="+str(le_mois)+" "+requete
            cr.execute("select count(*) from suivi_production_ta_cdt a inner join production_base_cdt b on a.base_id=b.id {0}".format(requete_week))
            for res_test in cr.fetchall():
                if res_test[0]>0:
                    feuille.set_row(0, 40)
                    compt=5
                    while compt<100:
                        feuille.set_row(compt, 11.25)
                        compt+=1
                    feuille.set_column('A:A', 2)
                    feuille.set_column('B:C', 30)
                    feuille.write('B4','NOM',style)    
                    feuille.write('C4','PRENOM',style) 
                    feuille.write('C1','NOM DOSSIER',style)
                    feuille.write('C2','TYPE DOSSIER',style)
                    feuille.write('C3','TARIF',style)
                    x=3
                    y=1
                    cr.execute("""select distinct a.base_id,b.name,c.name,b.tarif from suivi_production_ta_cdt a
                                    inner join production_base_cdt b on a.base_id=b.id
                                    left join production_base_type_cdt c on c.id=b.type_id
                                    {0} order by b.name asc""".format(requete_week))
                    for res in cr.fetchall():
                        feuille.merge_range(self.get_nom_colonne(x)+str(y)+':'+self.get_nom_colonne(x+1)+str(y),res[1],style)
                        feuille.merge_range(self.get_nom_colonne(x)+str(y+1)+':'+self.get_nom_colonne(x+1)+str(y+1),res[2],style)
                        feuille.merge_range(self.get_nom_colonne(x)+str(y+2)+':'+self.get_nom_colonne(x+1)+str(y+2),res[3],style)
                        feuille.set_column(self.get_nom_colonne(x)+':'+self.get_nom_colonne(x+1), 4.71)
                        feuille.write(self.get_nom_colonne(x)+str(y+3),'H PRD',style_yellow)
                        feuille.write(self.get_nom_colonne(x+1)+str(y+3),'CU',style_blue)
                        x=x+2
                    feuille.set_column(self.get_nom_colonne(x)+':'+self.get_nom_colonne(x), 10.29)
                    feuille.set_column(self.get_nom_colonne(x+1)+':'+self.get_nom_colonne(x+1), 5)
                    feuille.set_column(self.get_nom_colonne(x+2)+':'+self.get_nom_colonne(x+2), 8)
                    feuille.set_column(self.get_nom_colonne(x+3)+':'+self.get_nom_colonne(x+15), 16)
                        
                    feuille.write(self.get_nom_colonne(x)+str(y+3),'HEURES PROD',style_yellow)
                    feuille.write(self.get_nom_colonne(x+1)+str(y+3),'CA',style)
                    feuille.write(self.get_nom_colonne(x+2)+str(y+3),'CUMUL CU',style_blue)
                    feuille.write(self.get_nom_colonne(x+3)+str(y+3),'CA / H HORS RELANCE ET PANNES',style)
                    feuille.write(self.get_nom_colonne(x+4)+str(y+3),'CADENCE HORS RELANCE ET PANNES',style)
                    feuille.write(self.get_nom_colonne(x+5)+str(y+3),'HEURES PANNE',style)
                    feuille.write(self.get_nom_colonne(x+6)+str(y+3),'HEURES RELANCE',style)
                    feuille.write(self.get_nom_colonne(x+7)+str(y+3),'TOTAL HEURES REMUNEREES',style)
                    feuille.write(self.get_nom_colonne(x+8)+str(y+3),'TAUX PANNE',style)
                    feuille.write(self.get_nom_colonne(x+9)+str(y+3),'CA / H RELANCE INCLUSE',style)
                    feuille.write(self.get_nom_colonne(x+10)+str(y+3),'CADENCE RELANCE INCLUSE',style)
                    feuille.write(self.get_nom_colonne(x+11)+str(y+3),'CA / H RELANCE ET PANNE INCLUSE',style)
                    feuille.write(self.get_nom_colonne(x+12)+str(y+3),'CADENCE RELANCE ET PANNE INCLUSE',style)
                    feuille.write(self.get_nom_colonne(x+14)+str(y+3),'H NON REMUNEREES (20 h FORMATION)',style)
                    feuille.write(self.get_nom_colonne(x+15)+str(y+3),'TOTAL HEURE',style)
                    ##################### data
                    y_deb=y+4
                    y=y_deb
                    x=0
                    cr.execute("""select row_number() OVER (),name_related,prenom from 
                                    (select distinct name_related,prenom from suivi_production_ta_cdt a
                                    inner join production_base_cdt c on a.base_id=c.id
                                    left join hr_employee b on a.employee_id=b.id 
                                    {0} and employee_id is not null order by name_related,prenom) as t""".format(requete_week))
                    for res in cr.fetchall():
                        feuille.write(self.get_nom_colonne(x)+str(y),res[0],style) 
                        feuille.write(self.get_nom_colonne(x+1)+str(y),res[1],style) 
                        feuille.write(self.get_nom_colonne(x+2)+str(y),res[2],style) 
                        y=y+1
                    feuille.write(self.get_nom_colonne(x+2)+str(y+1),'CA DOSSIER',style) 
                    feuille.write(self.get_nom_colonne(x+2)+str(y+2),'CA / H DOSSIER',style) 
                    feuille.write(self.get_nom_colonne(x+2)+str(y+3),'CU / H DOSSIER',style) 
                    ##########################
                    x=3
                    y=y_deb
                    cr.execute("""select distinct employee_id,name_related,prenom 
                                    from suivi_production_ta_cdt a left join hr_employee b on a.employee_id=b.id
                                    inner join production_base_cdt c on a.base_id=c.id
                                    {0} and employee_id is not null   order by name_related,prenom""".format(requete_week))
                    for res in cr.fetchall():
                        employee_id=res[0]
                        h_prod_text=ca_text=cu_text=''
                        cr.execute("""select a.base_id,b.name,b.type_tarif
                                    from suivi_production_ta_cdt a
                                    inner join production_base_cdt b on a.base_id=b.id
                                    {0} group by a.base_id,b.name,b.type_tarif
                                    order by b.name asc""".format(requete_week))
                        for res2 in cr.fetchall():
                            base_id = res2[0]
                            type_tarif=res2[2]
                            h_prod=cu=''
                            cr.execute("""select sum(case when h_prod_jour is null then 0 else h_prod_jour end) as h_prod,sum(cu) as invitation
                                        from suivi_production_ta_cdt a
                                        inner join production_base_cdt c on a.base_id=c.id
                                        {0} and employee_id={1} and base_id={2}""".format(requete_week, employee_id, base_id))
                            for res1 in cr.fetchall():
                                h_prod=res1[0]
                                cu=res1[1]
                            h_prod_text+='+'+self.get_nom_colonne(x)+str(y)
                            if type_tarif=='invitation' :
                                ca_text+='+($'+self.get_nom_colonne(x)+'$3*'+self.get_nom_colonne(x+1)+str(y)+')'
                            if type_tarif=='heure' :
                                ca_text+='+($'+self.get_nom_colonne(x)+'$3*'+self.get_nom_colonne(x)+str(y)+')'
                            cu_text+='+'+self.get_nom_colonne(x+1)+str(y)
                            feuille.write(self.get_nom_colonne(x)+str(y),h_prod,style_yellow)
                            feuille.write(self.get_nom_colonne(x+1)+str(y),cu,style_blue)
                            x=x+2
                        cr.execute("""select sum(case when h_panne is null then 0 else h_panne end) as h_panne,
                                        sum(case when h_formation is null then 0 else h_formation end) as h_formation,
                                        sum(case when h_relance is null then 0 else h_relance end) as h_relance
                                        from suivi_production_ta_cdt a
                                        inner join production_base_cdt b on a.base_id=b.id
                                        {0} and employee_id={1}""".format(requete_week, employee_id))
                        for res1 in cr.fetchall():
                            h_panne=res1[0]
                            h_formation=res1[1]
                            h_relance=res1[2]
                        feuille.write_formula(self.get_nom_colonne(x)+str(y),h_prod_text[1:],style_yellow)
                        feuille.write_formula(self.get_nom_colonne(x+1)+str(y),ca_text[1:],style)
                        feuille.write_formula(self.get_nom_colonne(x+2)+str(y),cu_text[1:],style_blue)
                        feuille.write_formula(self.get_nom_colonne(x+3)+str(y),self.get_nom_colonne(x+1)+str(y)+'/'+self.get_nom_colonne(x)+str(y),style_number)
                        feuille.write_formula(self.get_nom_colonne(x+4)+str(y),self.get_nom_colonne(x+2)+str(y)+'/'+self.get_nom_colonne(x)+str(y),style_number)
                        feuille.write(self.get_nom_colonne(x+5)+str(y),h_panne,style)
                        feuille.write(self.get_nom_colonne(x+6)+str(y),h_relance,style)
                        feuille.write_formula(self.get_nom_colonne(x+7)+str(y),self.get_nom_colonne(x)+str(y)+'+'+self.get_nom_colonne(x+5)+str(y)+'+'+self.get_nom_colonne(x+6)+str(y),style_number)
                        feuille.write_formula(self.get_nom_colonne(x+8)+str(y),self.get_nom_colonne(x+5)+str(y)+'/'+self.get_nom_colonne(x+7)+str(y),style_pourcentage)
                        feuille.write_formula(self.get_nom_colonne(x+9)+str(y),self.get_nom_colonne(x+1)+str(y)+'/('+self.get_nom_colonne(x)+str(y)+'+'+self.get_nom_colonne(x+6)+str(y)+')',style_number)
                        feuille.write_formula(self.get_nom_colonne(x+10)+str(y),self.get_nom_colonne(x+2)+str(y)+'/('+self.get_nom_colonne(x)+str(y)+'+'+self.get_nom_colonne(x+6)+str(y)+')',style_number)
                        feuille.write_formula(self.get_nom_colonne(x+11)+str(y),self.get_nom_colonne(x+1)+str(y)+'/'+self.get_nom_colonne(x+7)+str(y),style_number)
                        feuille.write_formula(self.get_nom_colonne(x+12)+str(y),self.get_nom_colonne(x+2)+str(y)+'/'+self.get_nom_colonne(x+7)+str(y),style_number)
                        feuille.write(self.get_nom_colonne(x+14)+str(y),h_formation,style)
                        feuille.write_formula(self.get_nom_colonne(x+15)+str(y),self.get_nom_colonne(x+14)+str(y)+'+'+self.get_nom_colonne(x+7)+str(y),style_number)
                    
                        y=y+1
                        x=3
                    cr.execute("""select a.base_id,(select count(distinct employee_id) from suivi_production_ta_cdt a inner join production_base_cdt c on a.base_id=c.id {0} and employee_id is not null),b.type_tarif,b.name
                                    from suivi_production_ta_cdt a
                                    inner join production_base_cdt b on a.base_id=b.id
                                    {0} group by a.base_id,b.type_tarif,b.name order by b.name""".format(requete_week))
                    for res in cr.fetchall():
                        nb=res[1]
                        type_tarif=res[2]
                        feuille.write_formula(self.get_nom_colonne(x)+str(y),'SUM('+self.get_nom_colonne(x)+str(y_deb)+':'+self.get_nom_colonne(x)+str(y_deb+res[1]-1)+')',style_yellow)
                        feuille.write_formula(self.get_nom_colonne(x+1)+str(y),'SUM('+self.get_nom_colonne(x+1)+str(y_deb)+':'+self.get_nom_colonne(x+1)+str(y_deb+res[1]-1)+')',style_blue)
                        
                        feuille.merge_range(self.get_nom_colonne(x)+str(y+1)+':'+self.get_nom_colonne(x+1)+str(y+1),'',style)
                        feuille.merge_range(self.get_nom_colonne(x)+str(y+2)+':'+self.get_nom_colonne(x+1)+str(y+2),'',style)
                        feuille.merge_range(self.get_nom_colonne(x)+str(y+3)+':'+self.get_nom_colonne(x+1)+str(y+3),'',style)
                        if type_tarif=='invitation' :
                            feuille.write_formula(self.get_nom_colonne(x)+str(y+1),self.get_nom_colonne(x)+'3*'+str(self.get_nom_colonne(x+1)+str(y)),style_number)
                        if type_tarif=='heure' :
                            feuille.write_formula(self.get_nom_colonne(x)+str(y+1),self.get_nom_colonne(x)+'3*'+str(self.get_nom_colonne(x)+str(y)),style_number)
                        if type_tarif == 'retour':
                            feuille.write_formula(self.get_nom_colonne(x) + str(y + 1), self.get_nom_colonne(x) + '3*0', style_number)

                        feuille.write_formula(self.get_nom_colonne(x)+str(y+2),self.get_nom_colonne(x)+str(y+1)+'/'+self.get_nom_colonne(x)+str(y),style_number)
                        feuille.write_formula(self.get_nom_colonne(x)+str(y+3),self.get_nom_colonne(x+1)+str(y)+'/'+self.get_nom_colonne(x)+str(y),style_number)
                        
                        x=x+2
                    feuille.write_formula(self.get_nom_colonne(x)+str(y),'SUM('+self.get_nom_colonne(x)+str(y_deb)+':'+self.get_nom_colonne(x)+str(y_deb+nb-1)+')',style_yellow)
                    feuille.write_formula(self.get_nom_colonne(x+1)+str(y),'SUM('+self.get_nom_colonne(x+1)+str(y_deb)+':'+self.get_nom_colonne(x+1)+str(y_deb+nb-1)+')',style)
                    feuille.write_formula(self.get_nom_colonne(x+2)+str(y),'SUM('+self.get_nom_colonne(x+2)+str(y_deb)+':'+self.get_nom_colonne(x+2)+str(y_deb+nb-1)+')',style_blue)
                    feuille.write_formula(self.get_nom_colonne(x+3)+str(y),self.get_nom_colonne(x+1)+str(y)+'/'+self.get_nom_colonne(x)+str(y),style_number)
                    feuille.write_formula(self.get_nom_colonne(x+4)+str(y),self.get_nom_colonne(x+2)+str(y)+'/'+self.get_nom_colonne(x)+str(y),style_number)
                    feuille.write_formula(self.get_nom_colonne(x+5)+str(y),'SUM('+self.get_nom_colonne(x+5)+str(y_deb)+':'+self.get_nom_colonne(x+5)+str(y_deb+nb-1)+')',style)
                    feuille.write_formula(self.get_nom_colonne(x+6)+str(y),'SUM('+self.get_nom_colonne(x+6)+str(y_deb)+':'+self.get_nom_colonne(x+6)+str(y_deb+nb-1)+')',style)
                    feuille.write_formula(self.get_nom_colonne(x+7)+str(y),self.get_nom_colonne(x)+str(y)+'+'+self.get_nom_colonne(x+5)+str(y)+'+'+self.get_nom_colonne(x+6)+str(y),style_number)
                    feuille.write_formula(self.get_nom_colonne(x+8)+str(y),self.get_nom_colonne(x+5)+str(y)+'/'+self.get_nom_colonne(x+7)+str(y),style_pourcentage)
                    feuille.write_formula(self.get_nom_colonne(x+9)+str(y),self.get_nom_colonne(x+1)+str(y)+'/('+self.get_nom_colonne(x)+str(y)+'+'+self.get_nom_colonne(x+6)+str(y)+')',style_number)
                    feuille.write_formula(self.get_nom_colonne(x+10)+str(y),self.get_nom_colonne(x+2)+str(y)+'/('+self.get_nom_colonne(x)+str(y)+'+'+self.get_nom_colonne(x+6)+str(y)+')',style_number)
                    feuille.write_formula(self.get_nom_colonne(x+11)+str(y),self.get_nom_colonne(x+1)+str(y)+'/'+self.get_nom_colonne(x+7)+str(y),style_number)
                    feuille.write_formula(self.get_nom_colonne(x+12)+str(y),self.get_nom_colonne(x+2)+str(y)+'/'+self.get_nom_colonne(x+7)+str(y),style_number)
                    feuille.write_formula(self.get_nom_colonne(x+14)+str(y),'SUM('+self.get_nom_colonne(x+14)+str(y_deb)+':'+self.get_nom_colonne(x+14)+str(y_deb+nb-1)+')',style)
                    feuille.write_formula(self.get_nom_colonne(x+15)+str(y),self.get_nom_colonne(x+14)+str(y)+'+'+self.get_nom_colonne(x+7)+str(y),style_number)
                        
        feuille = workbook.add_worksheet(mois) 
        requete_month =" where extract(year from date)="+str(annee)+" and extract(month from date)="+str(le_mois)+" "+requete
        cr.execute("select count(*) from suivi_production_ta_cdt a inner join production_base_cdt b on a.base_id=b.id {0}".format(requete_month))
        for res_test in cr.fetchall():
            
            if res_test[0]>0:
                feuille.set_row(0, 40)
                compt=5
                while compt<100:
                    feuille.set_row(compt, 11.25)
                    compt+=1
                feuille.set_column('A:A', 2)
                feuille.set_column('B:C', 30)
                feuille.write('B4','NOM',style)    
                feuille.write('C4','PRENOM',style) 
                feuille.write('C1','NOM DOSSIER',style)
                feuille.write('C2','TYPE DOSSIER',style)
                feuille.write('C3','TARIF',style)
                x=3
                y=1
                cr.execute("""select distinct a.base_id,b.name,c.name,b.tarif from suivi_production_ta_cdt a
                                inner join production_base_cdt b on a.base_id=b.id
                                left join production_base_type_cdt c on c.id=b.type_id
                                {0} order by b.name""".format(requete_month))
                for res in cr.fetchall():
                    feuille.merge_range(self.get_nom_colonne(x)+str(y)+':'+self.get_nom_colonne(x+1)+str(y),res[1],style)
                    feuille.merge_range(self.get_nom_colonne(x)+str(y+1)+':'+self.get_nom_colonne(x+1)+str(y+1),res[2],style)
                    feuille.merge_range(self.get_nom_colonne(x)+str(y+2)+':'+self.get_nom_colonne(x+1)+str(y+2),res[3],style)
                    feuille.set_column(self.get_nom_colonne(x)+':'+self.get_nom_colonne(x+1), 4.71)
                    feuille.write(self.get_nom_colonne(x)+str(y+3),'H PRD',style_yellow)
                    feuille.write(self.get_nom_colonne(x+1)+str(y+3),'CU',style_blue)
                    x=x+2
                feuille.set_column(self.get_nom_colonne(x)+':'+self.get_nom_colonne(x), 10.29)
                feuille.set_column(self.get_nom_colonne(x+1)+':'+self.get_nom_colonne(x+1), 5)
                feuille.set_column(self.get_nom_colonne(x+2)+':'+self.get_nom_colonne(x+2), 8)
                feuille.set_column(self.get_nom_colonne(x+3)+':'+self.get_nom_colonne(x+15), 16)
                    
                feuille.write(self.get_nom_colonne(x)+str(y+3),'HEURES PROD',style_yellow)
                feuille.write(self.get_nom_colonne(x+1)+str(y+3),'CA',style)
                feuille.write(self.get_nom_colonne(x+2)+str(y+3),'CUMUL CU',style_blue)
                feuille.write(self.get_nom_colonne(x+3)+str(y+3),'CA / H HORS RELANCE ET PANNES',style)
                feuille.write(self.get_nom_colonne(x+4)+str(y+3),'CADENCE HORS RELANCE ET PANNES',style)
                feuille.write(self.get_nom_colonne(x+5)+str(y+3),'HEURES PANNE',style)
                feuille.write(self.get_nom_colonne(x+6)+str(y+3),'HEURES RELANCE',style)
                feuille.write(self.get_nom_colonne(x+7)+str(y+3),'TOTAL HEURES REMUNEREES',style)
                feuille.write(self.get_nom_colonne(x+8)+str(y+3),'TAUX PANNE',style)
                feuille.write(self.get_nom_colonne(x+9)+str(y+3),'CA / H RELANCE INCLUSE',style)
                feuille.write(self.get_nom_colonne(x+10)+str(y+3),'CADENCE RELANCE INCLUSE',style)
                feuille.write(self.get_nom_colonne(x+11)+str(y+3),'CA / H RELANCE ET PANNE INCLUSE',style)
                feuille.write(self.get_nom_colonne(x+12)+str(y+3),'CADENCE RELANCE ET PANNE INCLUSE',style)
                feuille.write(self.get_nom_colonne(x+14)+str(y+3),'H NON REMUNEREES (20 h FORMATION)',style)
                feuille.write(self.get_nom_colonne(x+15)+str(y+3),'TOTAL HEURE',style)
                ##################### data
                y_deb=y+4
                y=y_deb
                x=0
                cr.execute("""select row_number() OVER (),name_related,prenom from 
                                (select distinct name_related,prenom from suivi_production_ta_cdt a
                                left join hr_employee b on a.employee_id=b.id 
                                inner join production_base_cdt c on a.base_id=c.id
                                {0} and employee_id is not null order by name_related,prenom) as t""".format(requete_month))
                for res in cr.fetchall():
                    feuille.write(self.get_nom_colonne(x)+str(y),res[0],style) 
                    feuille.write(self.get_nom_colonne(x+1)+str(y),res[1],style) 
                    feuille.write(self.get_nom_colonne(x+2)+str(y),res[2],style) 
                    y=y+1
                feuille.write(self.get_nom_colonne(x+2)+str(y+1),'CA DOSSIER',style) 
                feuille.write(self.get_nom_colonne(x+2)+str(y+2),'CA / H DOSSIER',style) 
                feuille.write(self.get_nom_colonne(x+2)+str(y+3),'CU / H DOSSIER',style) 
                ##########################
                x=3
                y=y_deb
                cr.execute("""select distinct employee_id,name_related,prenom 
                            from suivi_production_ta_cdt a left join hr_employee b on a.employee_id=b.id
                            inner join production_base_cdt c on a.base_id=c.id
                            {0} and employee_id is not null   order by name_related,prenom""".format(requete_month))
                for res in cr.fetchall():
                    employee_id=res[0]
                    h_prod_text=ca_text=cu_text=''
                    cr.execute("""select a.base_id,b.name,b.type_tarif
                                    from suivi_production_ta_cdt a
                                    inner join production_base_cdt b on a.base_id=b.id
                                    {0} group by a.base_id,b.name,b.type_tarif order by b.name asc""".format(requete_month))
                    for res2 in cr.fetchall():
                        base_id = res2[0]
                        type_tarif=res2[2]
                        h_prod=cu=''
                        cr.execute("""select sum(case when h_prod_jour is null then 0 else h_prod_jour end) as h_prod,sum(cu) as invitation
                                        from suivi_production_ta_cdt a
                                        inner join production_base_cdt c on a.base_id=c.id
                                        {0} and employee_id={1} and base_id={2}""".format(requete_month, employee_id, base_id))
                        for res1 in cr.fetchall():
                            h_prod=res1[0]
                            cu=res1[1]
                        retour = 0
                        cr.execute("""select sum(retour) as invitation
                                        from production_retour_cdt a
                                        inner join production_base_cdt b on b.name=a.base
                                        inner join production_login_cdt c on c.name=a.login
                                        where year={0}
                                        and month={1}
                                        and b.id  = {3}
                                        and c.employee_id = {2}""".format(annee, le_mois,employee_id, base_id))
                        for res1 in cr.fetchall():
                            if res1[0]:
                                retour=res1[0]
                        h_prod_text+='+'+self.get_nom_colonne(x)+str(y)
                        if type_tarif=='invitation':
                            ca_text+='+($'+self.get_nom_colonne(x)+'$3*'+self.get_nom_colonne(x+1)+str(y)+')'
                        if type_tarif=='heure':
                            ca_text+='+($'+self.get_nom_colonne(x)+'$3*'+self.get_nom_colonne(x)+str(y)+')'
                        if type_tarif=='retour':
                            ca_text+='+($'+self.get_nom_colonne(x)+'$3*'+str(retour)+')'
                        cu_text+='+'+self.get_nom_colonne(x+1)+str(y)
                        feuille.write(self.get_nom_colonne(x)+str(y),h_prod,style_yellow)
                        feuille.write(self.get_nom_colonne(x+1)+str(y),cu,style_blue)
                        x=x+2
                    cr.execute("""select sum(case when h_panne is null then 0 else h_panne end) as h_panne,
                                    sum(case when h_formation is null then 0 else h_formation end) as h_formation,
                                    sum(case when h_relance is null then 0 else h_relance end) as h_relance
                                    from suivi_production_ta_cdt a
                                    inner join production_base_cdt b on a.base_id=b.id
                                    {0} and employee_id={1}""".format(requete_month, employee_id))
                    for res1 in cr.fetchall():
                        h_panne=res1[0]
                        h_formation=res1[1]
                        h_relance=res1[2]
                    feuille.write_formula(self.get_nom_colonne(x)+str(y),h_prod_text[1:],style_yellow)
                    feuille.write_formula(self.get_nom_colonne(x+1)+str(y),ca_text[1:],style)
                    feuille.write_formula(self.get_nom_colonne(x+2)+str(y),cu_text[1:],style_blue)
                    feuille.write_formula(self.get_nom_colonne(x+3)+str(y),self.get_nom_colonne(x+1)+str(y)+'/'+self.get_nom_colonne(x)+str(y),style_number)
                    feuille.write_formula(self.get_nom_colonne(x+4)+str(y),self.get_nom_colonne(x+2)+str(y)+'/'+self.get_nom_colonne(x)+str(y),style_number)
                    feuille.write(self.get_nom_colonne(x+5)+str(y),h_panne,style)
                    feuille.write(self.get_nom_colonne(x+6)+str(y),h_relance,style)
                    feuille.write_formula(self.get_nom_colonne(x+7)+str(y),self.get_nom_colonne(x)+str(y)+'+'+self.get_nom_colonne(x+5)+str(y)+'+'+self.get_nom_colonne(x+6)+str(y),style_number)
                    feuille.write_formula(self.get_nom_colonne(x+8)+str(y),self.get_nom_colonne(x+5)+str(y)+'/'+self.get_nom_colonne(x+7)+str(y),style_pourcentage)
                    feuille.write_formula(self.get_nom_colonne(x+9)+str(y),self.get_nom_colonne(x+1)+str(y)+'/('+self.get_nom_colonne(x)+str(y)+'+'+self.get_nom_colonne(x+6)+str(y)+')',style_number)
                    feuille.write_formula(self.get_nom_colonne(x+10)+str(y),self.get_nom_colonne(x+2)+str(y)+'/('+self.get_nom_colonne(x)+str(y)+'+'+self.get_nom_colonne(x+6)+str(y)+')',style_number)
                    feuille.write_formula(self.get_nom_colonne(x+11)+str(y),self.get_nom_colonne(x+1)+str(y)+'/'+self.get_nom_colonne(x+7)+str(y),style_number)
                    feuille.write_formula(self.get_nom_colonne(x+12)+str(y),self.get_nom_colonne(x+2)+str(y)+'/'+self.get_nom_colonne(x+7)+str(y),style_number)
                    feuille.write(self.get_nom_colonne(x+14)+str(y),h_formation,style)
                    feuille.write_formula(self.get_nom_colonne(x+15)+str(y),self.get_nom_colonne(x+14)+str(y)+'+'+self.get_nom_colonne(x+7)+str(y),style_number)
                
                    y=y+1
                    x=3
                cr.execute("""select a.base_id,(select count(distinct employee_id) from suivi_production_ta_cdt a inner join production_base_cdt c on a.base_id=c.id {0} and employee_id is not null),b.type_tarif,b.name
                                from suivi_production_ta_cdt a
                                inner join production_base_cdt b on a.base_id=b.id
                                {0} group by a.base_id,b.type_tarif,b.name order by b.name""".format(requete_month))
                for res in cr.fetchall():
                    base_id = res[0]
                    nb=res[1]
                    type_tarif=res[2]
                    feuille.write_formula(self.get_nom_colonne(x)+str(y),'SUM('+self.get_nom_colonne(x)+str(y_deb)+':'+self.get_nom_colonne(x)+str(y_deb+res[1]-1)+')',style_yellow)
                    feuille.write_formula(self.get_nom_colonne(x+1)+str(y),'SUM('+self.get_nom_colonne(x+1)+str(y_deb)+':'+self.get_nom_colonne(x+1)+str(y_deb+res[1]-1)+')',style_blue)
                    
                    feuille.merge_range(self.get_nom_colonne(x)+str(y+1)+':'+self.get_nom_colonne(x+1)+str(y+1),'',style)
                    feuille.merge_range(self.get_nom_colonne(x)+str(y+2)+':'+self.get_nom_colonne(x+1)+str(y+2),'',style)
                    feuille.merge_range(self.get_nom_colonne(x)+str(y+3)+':'+self.get_nom_colonne(x+1)+str(y+3),'',style)
                    retour = 0
                    cr.execute("""select sum(retour) as invitation
                                                            from production_retour_cdt d
                                                            inner join production_base_cdt b on b.name=d.base
                                                            inner join production_login_cdt c on c.name=d.login
                                                            inner join hr_employee a on a.id=c.employee_id
                                                            where year={0}
                                                            and month={1}
                                                            and b.id  = {2} {3}""".format(annee, le_mois, base_id, requete))
                    for res1 in cr.fetchall():
                        retour = res1[0]
                    if type_tarif=='invitation':
                        feuille.write_formula(self.get_nom_colonne(x)+str(y+1),self.get_nom_colonne(x)+'3*'+str(self.get_nom_colonne(x+1)+str(y)),style_number)
                    if type_tarif=='heure':
                        feuille.write_formula(self.get_nom_colonne(x)+str(y+1),self.get_nom_colonne(x)+'3*'+str(self.get_nom_colonne(x)+str(y)),style_number)
                    if type_tarif == 'retour':
                        feuille.write_formula(self.get_nom_colonne(x) + str(y + 1), self.get_nom_colonne(x) + '3*' + str(retour), style_number)
                    feuille.write_formula(self.get_nom_colonne(x)+str(y+2),self.get_nom_colonne(x)+str(y+1)+'/'+self.get_nom_colonne(x)+str(y),style_number)
                    feuille.write_formula(self.get_nom_colonne(x)+str(y+3),self.get_nom_colonne(x+1)+str(y)+'/'+self.get_nom_colonne(x)+str(y),style_number)
                    
                    x=x+2
                feuille.write_formula(self.get_nom_colonne(x)+str(y),'SUM('+self.get_nom_colonne(x)+str(y_deb)+':'+self.get_nom_colonne(x)+str(y_deb+nb-1)+')',style_yellow)
                feuille.write_formula(self.get_nom_colonne(x+1)+str(y),'SUM('+self.get_nom_colonne(x+1)+str(y_deb)+':'+self.get_nom_colonne(x+1)+str(y_deb+nb-1)+')',style)
                feuille.write_formula(self.get_nom_colonne(x+2)+str(y),'SUM('+self.get_nom_colonne(x+2)+str(y_deb)+':'+self.get_nom_colonne(x+2)+str(y_deb+nb-1)+')',style_blue)
                feuille.write_formula(self.get_nom_colonne(x+3)+str(y),self.get_nom_colonne(x+1)+str(y)+'/'+self.get_nom_colonne(x)+str(y),style_number)
                feuille.write_formula(self.get_nom_colonne(x+4)+str(y),self.get_nom_colonne(x+2)+str(y)+'/'+self.get_nom_colonne(x)+str(y),style_number)
                feuille.write_formula(self.get_nom_colonne(x+5)+str(y),'SUM('+self.get_nom_colonne(x+5)+str(y_deb)+':'+self.get_nom_colonne(x+5)+str(y_deb+nb-1)+')',style)
                feuille.write_formula(self.get_nom_colonne(x+6)+str(y),'SUM('+self.get_nom_colonne(x+6)+str(y_deb)+':'+self.get_nom_colonne(x+6)+str(y_deb+nb-1)+')',style)
                feuille.write_formula(self.get_nom_colonne(x+7)+str(y),self.get_nom_colonne(x)+str(y)+'+'+self.get_nom_colonne(x+5)+str(y)+'+'+self.get_nom_colonne(x+6)+str(y),style_number)
                feuille.write_formula(self.get_nom_colonne(x+8)+str(y),self.get_nom_colonne(x+5)+str(y)+'/'+self.get_nom_colonne(x+7)+str(y),style_pourcentage)
                feuille.write_formula(self.get_nom_colonne(x+9)+str(y),self.get_nom_colonne(x+1)+str(y)+'/('+self.get_nom_colonne(x)+str(y)+'+'+self.get_nom_colonne(x+6)+str(y)+')',style_number)
                feuille.write_formula(self.get_nom_colonne(x+10)+str(y),self.get_nom_colonne(x+2)+str(y)+'/('+self.get_nom_colonne(x)+str(y)+'+'+self.get_nom_colonne(x+6)+str(y)+')',style_number)
                feuille.write_formula(self.get_nom_colonne(x+11)+str(y),self.get_nom_colonne(x+1)+str(y)+'/'+self.get_nom_colonne(x+7)+str(y),style_number)
                feuille.write_formula(self.get_nom_colonne(x+12)+str(y),self.get_nom_colonne(x+2)+str(y)+'/'+self.get_nom_colonne(x+7)+str(y),style_number)
                feuille.write_formula(self.get_nom_colonne(x+14)+str(y),'SUM('+self.get_nom_colonne(x+14)+str(y_deb)+':'+self.get_nom_colonne(x+14)+str(y_deb+nb-1)+')',style)
                feuille.write_formula(self.get_nom_colonne(x+15)+str(y),self.get_nom_colonne(x+14)+str(y)+'+'+self.get_nom_colonne(x+7)+str(y),style_number)
                    
        workbook.close()
        return fichier
    
    def generer_fh_cdt(self, cr, uid,ids=True, context=None):
        reload(sys)
        sys.setdefaultencoding("UTF8")       
        for obj in self.read(cr, uid, ids, ['annee','mois','superviseur_id'], context=context):
            mois=obj['mois']
            annee=obj['annee']
            #superviseur_id=obj['superviseur_id'][0]
        fichier=self.generer_fh_cdt_(cr, uid,mois,annee, ids, context)
        return self.get_return(cr, uid,fichier, ids, context)
    def generer_fh_cdt_(self, cr, uid,mois,annee,ids=True, context=None):
        reload(sys)
        sys.setdefaultencoding("UTF8")   
        le_mois=self.get_le_mois(mois)
        superviseurs=''
        
        ###################### Export Excel Trame CE###################
        fichier="CDT-Feuille d'heures "+mois+" "+str(annee)+".xlsx"
        workbook = xlsxwriter.Workbook(sortie + fichier)
        style_x = workbook.add_format({   'text_wrap' : True,'bold' :1,  'valign' : 'vcenter','bg_color' : 'white'})
        style1_heure = workbook.add_format({ 'align' : 'center','valign' : 'vcenter','right' : 2,'left' :1,'top':1,'bottom':1,'num_format' : 'hh:mm'})
        style = workbook.add_format({   'text_wrap' : True,'bold' :1,  'align' : 'center','valign' : 'vcenter','border' : 2})
        style1_vert_nombre = workbook.add_format({ 'num_format' : '#,##0.00', 'align' : 'center','valign' : 'vcenter','right' : 2,'left' :1,'top':1,'bottom':1,'bg_color' : '#FFFF99'})
        style1_rouge = workbook.add_format({ 'align' : 'center','valign' : 'vcenter','right' : 2,'left' :1,'top':1,'bottom':1,'bg_color' : 'red'})
        style1_conge = workbook.add_format({ 'align' : 'center','valign' : 'vcenter','right' : 2,'left' :1,'top':1,'bottom':1,'bg_color' : '#92D050'})
        style1_chibi = workbook.add_format({ 'align' : 'center','valign' : 'vcenter','right' : 2,'left' :1,'top':1,'bottom':1,'bg_color' : '#C3F3FD'})
        style1_presence_nombre = workbook.add_format({'num_format' : '#,##0.00', 'align' : 'center','valign' : 'vcenter','right' : 2,'left' :1,'top':1,'bottom':1,'bg_color' : '#FFCC99'})
        style1_bleu = workbook.add_format({ 'align' : 'center','valign' : 'vcenter','right' : 2,'left' :1,'top':1,'bottom':1,'bg_color' : 'blue'})
        style_nom_ta = workbook.add_format({ 'align' : 'center','valign' : 'vcenter','right' : 2,'left' :1,'top':1,'bottom':1,'bg_color' : '#CCC0DA','bold' :1})
        style_nom_sup = workbook.add_format({ 'align' : 'center','valign' : 'vcenter','right' : 2,'left' :1,'top':1,'bottom':1,'bg_color' : '#D7E4BC','bold' :1})
        style_nom_red = workbook.add_format({ 'align' : 'center','valign' : 'vcenter','right' : 2,'left' :1,'top':1,'bottom':1,'bg_color' : 'red','bold' :1})
        style_nom_white = workbook.add_format({ 'align' : 'center','valign' : 'vcenter','right' : 2,'left' :1,'top':1,'bottom':1,'bg_color' : 'white','bold' :1})
        style1 = workbook.add_format({ 'align' : 'center','valign' : 'vcenter','right' : 2,'left' :1,'top':1,'bottom':1})
        style1_black = workbook.add_format({ 'align' : 'center','valign' : 'vcenter','right' : 2,'left' :1,'top':1,'bottom':1,'bg_color' : '#000000'})
        style1_red = workbook.add_format({ 'align' : 'center','valign' : 'vcenter','right' : 2,'left' :1,'top':1,'bottom':1,'bg_color' : 'red'})
        style1_rotation = workbook.add_format({ 'align' : 'center','valign' : 'vcenter','right' : 2,'left' :1,'top':1,'bottom':1,'rotation' : 90})
        style_ss_h_tot_ta = workbook.add_format({ 'bold' :1, 'align' : 'center','valign' : 'vcenter','right' : 2,'left' :1,'top':1,'bottom':1,'bg_color' : '#FF99CC'})
        style_ss_h_tot_ta_nombre = workbook.add_format({'bold' :1,  'num_format' : '#,##0.00', 'align' : 'center','valign' : 'vcenter','right' : 2,'left' :1,'top':1,'bottom':1,'bg_color' : '#FF99CC'})
        style_ss_h_tot_enc = workbook.add_format({'bold' :1,  'align' : 'center','valign' : 'vcenter','right' : 2,'left' :1,'top':1,'bottom':1,'bg_color' : '#99CC00'})
        style_ss_h_tot_enc_nombre = workbook.add_format({'bold' :1,  'num_format' : '#,##0.00', 'align' : 'center','valign' : 'vcenter','right' : 2,'left' :1,'top':1,'bottom':1,'bg_color' : '#99CC00' })
        style_tot_gen = workbook.add_format({'bold' :1,  'align' : 'center','valign' : 'vcenter','right' : 2,'left' :1,'top':1,'bottom':1,'bg_color' : '#99CCFF' })
        style_tot_gen_nombre = workbook.add_format({ 'bold' :1, 'num_format' : '#,##0.00', 'align' : 'center','valign' : 'vcenter','right' : 2,'left' :1,'top':1,'bottom':1,'bg_color' : '#99CCFF'})
        cr.execute("select distinct date from hr_employee_presence_cdt\
                    where to_char(date,'MM')='"+str(le_mois)+"'\
                    and to_char(date,'YYYY')='"+str(annee)+"'\
                    order by date",(tuple(ids),))
        for res in cr.fetchall():
            date=res[0]
            feuille=workbook.add_worksheet(res[0])
            feuille.set_zoom(85)
            feuille.set_column('A:A', 42.14)
            feuille.set_column('B:E', 8.59)
            feuille.set_column('F:F', 9.43)
            feuille.set_column('G:G', 10.71)
            feuille.set_column('I:I', 8.86)
            feuille.set_column('J:J', 10.71)
            feuille.set_column('K:L', 5.86)
            feuille.set_column('M:M', 57.71)
            
            x=1
            feuille.merge_range('A'+str(x)+':L'+str(x),'Opération : CDT',style_x)
            feuille.merge_range('A'+str(x+1)+':M'+str(x+1),'Chef de Projet :  ',style_x)
            feuille.merge_range('A'+str(x+2)+':M'+str(x+2),'Superviseurs : ',style_x)
            feuille.merge_range('A'+str(x+3)+':M'+str(x+3),'',style_x)
            feuille.write('A'+str(x+4),'',style_x)
            feuille.merge_range('B'+str(x+4)+':C'+str(x+4),'MATIN',style)
            feuille.merge_range('D'+str(x+4)+':E'+str(x+4),'APRES-MIDI',style)
            feuille.merge_range('F'+str(x+4)+':L'+str(x+4),'TOTAL HEURES',style)
            feuille.write('A'+str(x+5),'Nom et Prénom',style)
            feuille.write('B'+str(x+5),'Arr.',style)
            feuille.write('C'+str(x+5),'Dép.',style)
            feuille.write('D'+str(x+5),'Arr.',style)
            feuille.write('E'+str(x+5),'Dép.',style)
            feuille.write('F'+str(x+5),'H. Prés. (Nb)',style)
            feuille.write('G'+str(x+5),'Form° Initiale',style)
            feuille.write('H'+str(x+5),'Brief TELEXCEL',style)
            feuille.write('I'+str(x+5),'Panne',style)
            feuille.write('J'+str(x+5),'H.RELANCE',style)
            feuille.write('K'+str(x+5),'H. Fact.',style)
            feuille.write('L'+str(x+5),'',style)
            feuille.write('M'+str(x+5),'Commentaires',style)
            x=x+6
            h_pres_tot=h_for_tot=h_brief_tot=h_panne_tot=h_fact_tot=h_pnc_tot=heure_tot=''
            #######################SOUS-TOTAL HEURES TA###############################
            h_pres=h_for=h_brief=h_panne=h_fact=h_pnc=heure=''
            cr.execute("select emp.complete_name,am_arr,am_dep,pm_arr,pm_dep,h_formation,h_brief,h_panne,pres.state,to_char(date,'dd/mm/yyyy')\
                        from hr_employee_presence_cdt pres left join hr_employee emp on emp.id=pres.employee_id\
                        where date='"+date+"'\
                        and pres.category_id=(select id from hr_employee_category where name='TA')\
                        order by emp.complete_name",(tuple(ids),))
            for res1 in cr.fetchall():
                h_pres+='+F'+str(x)
                h_for+='+G'+str(x)
                h_brief+='+H'+str(x)
                h_panne+='+I'+str(x)
                h_pnc+='+J'+str(x)
                h_fact+='+K'+str(x)
                heure+='+L'+str(x)
                h_pres_tot+='+F'+str(x)
                h_for_tot+='+G'+str(x)
                h_brief_tot+='+H'+str(x)
                h_panne_tot+='+I'+str(x)
                h_pnc_tot+='+J'+str(x)
                h_fact_tot+='+K'+str(x)
                heure_tot+='+L'+str(x)
                complete_name=res1[0]
                am_arr=res1[1]
                am_dep=res1[2]
                pm_arr=res1[3]
                pm_dep=res1[4]
                finitiale=res1[5]
                brief=res1[6]
                panne=res1[7]
                state=res1[8]
                date1=res1[9]
                feuille.write('A'+str(x),complete_name,style_nom_ta)
                if state=="absent":
                    feuille.write('B'+str(x),'',style1_rouge)
                    feuille.write('C'+str(x),'',style1_rouge)
                    feuille.write('D'+str(x),'',style1_rouge)
                    feuille.write('E'+str(x),'',style1_rouge)
                    feuille.write('G'+str(x),'',style1)
                    feuille.write('H'+str(x),'',style1)
                    feuille.write('I'+str(x),'',style1_rouge)
                    feuille.write('J'+str(x),'',style1_chibi)
                elif state=="en_conge":
                    feuille.write('B'+str(x),'',style1_conge)
                    feuille.write('C'+str(x),'',style1_conge)
                    feuille.write('D'+str(x),'',style1_conge)
                    feuille.write('E'+str(x),'',style1_conge)
                    feuille.write('G'+str(x),'',style1)
                    feuille.write('H'+str(x),'',style1)
                    feuille.write('I'+str(x),'',style1_rouge)
                    feuille.write('J'+str(x),'',style1_chibi)
                else :
                    feuille.write('B'+str(x),datetime.datetime.strptime(date1+' '+self.FLOAT_TO_HMS(am_arr), '%d/%m/%Y %H:%M:%S').time(),style1_heure)
                    feuille.write('C'+str(x),datetime.datetime.strptime(date1+' '+self.FLOAT_TO_HMS(am_dep), '%d/%m/%Y %H:%M:%S').time(),style1_heure)
                    feuille.write('D'+str(x),datetime.datetime.strptime(date1+' '+self.FLOAT_TO_HMS(pm_arr), '%d/%m/%Y %H:%M:%S').time(),style1_heure)
                    feuille.write('E'+str(x),datetime.datetime.strptime(date1+' '+self.FLOAT_TO_HMS(pm_dep), '%d/%m/%Y %H:%M:%S').time(),style1_heure)
                    feuille.write('G'+str(x),finitiale,style1)
                    feuille.write('H'+str(x),brief,style1)
                    feuille.write('I'+str(x),panne,style1_rouge)
                    feuille.write('J'+str(x),'',style1_chibi)
                if state=='present':
                    state=''
                feuille.write('F'+str(x),'=(SUM(C'+str(x)+'-B'+str(x)+')+(E'+str(x)+'-D'+str(x)+'))*24',style1_presence_nombre)
                feuille.write_formula('K'+str(x),'=F'+str(x)+'-G'+str(x)+'',style1_vert_nombre)
                feuille.write('L'+str(x),'',style1)
                feuille.write('M'+str(x),state,style1)
                x=x+1
            feuille.write('A'+str(x),'SOUS-TOTAL HEURES TA',style_ss_h_tot_ta)
            feuille.write('B'+str(x),'',style_ss_h_tot_ta)
            feuille.write('C'+str(x),'',style_ss_h_tot_ta)
            feuille.write('D'+str(x),'',style_ss_h_tot_ta)
            feuille.write('E'+str(x),'',style_ss_h_tot_ta)
            feuille.write_formula('F'+str(x),h_pres[1:],style_ss_h_tot_ta_nombre)
            feuille.write_formula('G'+str(x),h_for[1:],style_ss_h_tot_ta_nombre)
            feuille.write_formula('H'+str(x),h_brief[1:],style_ss_h_tot_ta_nombre)
            feuille.write_formula('I'+str(x),h_panne[1:],style_ss_h_tot_ta_nombre)
            feuille.write_formula('J'+str(x),h_pnc[1:],style_ss_h_tot_ta_nombre)
            feuille.write_formula('K'+str(x),h_fact[1:],style_ss_h_tot_ta_nombre)
            feuille.write_formula('L'+str(x),heure[1:],style_ss_h_tot_ta_nombre)
            feuille.write('M'+str(x),'',style1)
            x=x+1
            ##################SOUS-TOTAL HEURES Encadrants#########################
            h_pres=h_for=h_brief=h_panne=h_fact=h_pnc=heure=''
            cr.execute("select emp.complete_name,am_arr,am_dep,pm_arr,pm_dep,h_formation,h_brief,h_panne,pres.state,to_char(date,'dd/mm/yyyy') \
                        from hr_employee_presence_cdt pres left join hr_employee emp on emp.id=pres.employee_id\
                        where date='"+date+"'\
                        and pres.category_id in (select id from hr_employee_category where name in ('SUP','Chef de plateau'))\
                        order by emp.complete_name",(tuple(ids),))
            for res1 in cr.fetchall():
                h_pres+='+F'+str(x)
                h_for+='+G'+str(x)
                h_brief+='+H'+str(x)
                h_panne+='+I'+str(x)
                h_pnc+='+J'+str(x)
                h_fact+='+K'+str(x)
                heure+='+L'+str(x)
                h_pres_tot+='+F'+str(x)
                h_for_tot+='+G'+str(x)
                h_brief_tot+='+H'+str(x)
                h_panne_tot+='+I'+str(x)
                h_pnc_tot+='+J'+str(x)
                h_fact_tot+='+K'+str(x)
                heure_tot+='+L'+str(x)
                complete_name=res1[0]
                am_arr=res1[1]
                am_dep=res1[2]
                pm_arr=res1[3]
                pm_dep=res1[4]
                finitiale=res1[5]
                brief=res1[6]
                panne=res1[7]
                state=res1[8]
                date1=res1[9]
                feuille.write('A'+str(x),complete_name,style_nom_sup)
                if state=="absent":
                    feuille.write('B'+str(x),'',style1_bleu)
                    feuille.write('C'+str(x),'',style1_bleu)
                    feuille.write('D'+str(x),'',style1_bleu)
                    feuille.write('E'+str(x),'',style1_bleu)
                    feuille.write('G'+str(x),'',style1)
                    feuille.write('H'+str(x),'',style1)
                    feuille.write('I'+str(x),'',style1_rouge)
                    feuille.write('J'+str(x),'',style1_chibi)
                elif state=="en_conge":
                    feuille.write('B'+str(x),'',style1_conge)
                    feuille.write('C'+str(x),'',style1_conge)
                    feuille.write('D'+str(x),'',style1_conge)
                    feuille.write('E'+str(x),'',style1_conge)
                    feuille.write('G'+str(x),'',style1)
                    feuille.write('H'+str(x),'',style1)
                    feuille.write('I'+str(x),'',style1_rouge)
                    feuille.write('J'+str(x),'',style1_chibi)
                
                else :
                    feuille.write('B'+str(x),datetime.datetime.strptime(date1+' '+self.FLOAT_TO_HMS(am_arr), '%d/%m/%Y %H:%M:%S').time(),style1_heure)
                    feuille.write('C'+str(x),datetime.datetime.strptime(date1+' '+self.FLOAT_TO_HMS(am_dep), '%d/%m/%Y %H:%M:%S').time(),style1_heure)
                    feuille.write('D'+str(x),datetime.datetime.strptime(date1+' '+self.FLOAT_TO_HMS(pm_arr), '%d/%m/%Y %H:%M:%S').time(),style1_heure)
                    feuille.write('E'+str(x),datetime.datetime.strptime(date1+' '+self.FLOAT_TO_HMS(pm_dep), '%d/%m/%Y %H:%M:%S').time(),style1_heure)
                    feuille.write('G'+str(x),finitiale,style1)
                    feuille.write('H'+str(x),brief,style1)
                    feuille.write('I'+str(x),panne,style1_rouge)
                    feuille.write('J'+str(x),'',style1_chibi)
                if state=='present':
                    state=''
                feuille.write('F'+str(x),'=(SUM(C'+str(x)+'-B'+str(x)+')+(E'+str(x)+'-D'+str(x)+'))*24',style1_presence_nombre)
                feuille.write_formula('K'+str(x),'=F'+str(x)+'-G'+str(x)+'',style1_vert_nombre)
                feuille.write('L'+str(x),'',style1)
                feuille.write('M'+str(x),state,style1) 
                x=x+1 
            feuille.write('A'+str(x),'SOUS-TOTAL HEURES Encadrants',style_ss_h_tot_enc)
            feuille.write('B'+str(x),'',style_ss_h_tot_enc)
            feuille.write('C'+str(x),'',style_ss_h_tot_enc)
            feuille.write('D'+str(x),'',style_ss_h_tot_enc)
            feuille.write('E'+str(x),'',style_ss_h_tot_enc)
            feuille.write_formula('F'+str(x),h_pres[1:],style_ss_h_tot_enc_nombre)
            feuille.write_formula('G'+str(x),h_for[1:],style_ss_h_tot_enc_nombre)
            feuille.write_formula('H'+str(x),h_brief[1:],style_ss_h_tot_enc_nombre)
            feuille.write_formula('I'+str(x),h_panne[1:],style_ss_h_tot_enc_nombre)
            feuille.write_formula('J'+str(x),h_pnc[1:],style_ss_h_tot_enc_nombre)
            feuille.write_formula('K'+str(x),h_fact[1:],style_ss_h_tot_enc_nombre)
            feuille.write_formula('L'+str(x),heure[1:],style_ss_h_tot_enc_nombre)
            feuille.write('M'+str(x),'',style1)
            x=x+1
            ##############TOTAL GENERAL############################
            feuille.write('A'+str(x),'TOTAL GENERAL HEURES  ',style_tot_gen)
            feuille.write('B'+str(x),'',style_tot_gen)
            feuille.write('C'+str(x),'',style_tot_gen)
            feuille.write('D'+str(x),'',style_tot_gen)
            feuille.write('E'+str(x),'',style_tot_gen)
            feuille.write_formula('F'+str(x),h_pres_tot[1:],style_tot_gen_nombre)
            feuille.write_formula('G'+str(x),h_for_tot[1:],style_tot_gen_nombre)
            feuille.write_formula('H'+str(x),h_brief_tot[1:],style_tot_gen_nombre)
            feuille.write_formula('I'+str(x),h_panne_tot[1:],style_tot_gen_nombre)
            feuille.write_formula('J'+str(x),h_pnc_tot[1:],style_tot_gen_nombre)
            feuille.write_formula('K'+str(x),h_fact_tot[1:],style_tot_gen_nombre)
            feuille.write_formula('L'+str(x),heure_tot[1:],style_tot_gen_nombre)
            feuille.write('M'+str(x),'',style1)
        
        feuille_cumul=workbook.add_worksheet('Cumul')
        feuille_cumul.set_zoom(85)
        feuille_cumul.set_column('A:A', 42.14)
        feuille_cumul.set_column('B:E', 8.59)
        feuille_cumul.set_column('F:F', 9.43)
        feuille_cumul.set_column('G:G', 10.71)
        feuille_cumul.set_column('I:I', 8.86)
        feuille_cumul.set_column('J:J', 10.71)
        feuille_cumul.set_column('K:L', 5.86)
        feuille_cumul.set_column('M:M', 57.71)
        
        x=1
        feuille_cumul.merge_range('A'+str(x)+':L'+str(x),'Opération : CDT',style_x)
        feuille_cumul.merge_range('A'+str(x+1)+':M'+str(x+1),'Chef de Projet :  ',style_x)
        feuille_cumul.merge_range('A'+str(x+2)+':M'+str(x+2),'Superviseurs : ',style_x)
        feuille_cumul.merge_range('A'+str(x+3)+':M'+str(x+3),'',style_x)
        
        feuille_cumul.write('A'+str(x+5),'Nom et Prénom',style)
        feuille_cumul.write('B'+str(x+5),'H. Prés. (Nb)',style)
        feuille_cumul.write('C'+str(x+5),'Form° Initiale',style)
        feuille_cumul.write('D'+str(x+5),'Brief TELEXCEL',style)
        feuille_cumul.write('E'+str(x+5),'Panne',style)
        feuille_cumul.write('F'+str(x+5),'H.RELANCE',style)
        feuille_cumul.write('G'+str(x+5),'H. Fact.',style)
        x=x+6
        h_pres_tot=h_for_tot=h_brief_tot=h_panne_tot=h_fact_tot=h_relance_tot=heure_tot=''
        #######################SOUS-TOTAL HEURES TA###############################
        h_pres=h_for=h_brief=h_panne=h_fact=h_relance=heure=''
        cr.execute("select emp.complete_name,sum(case when h_pres is null then 0 else h_pres end),\
                    sum(case when h_formation is null then 0 else h_formation end),\
                    sum(case when h_brief is null then 0 else h_brief end),\
                    sum(case when h_panne is null then 0 else h_panne end),\
                    sum(case when h_relance is null then 0 else h_relance end)\
                    from hr_employee_presence_cdt pres left join hr_employee emp on emp.id=pres.employee_id\
                    where to_char(date,'MM')='"+str(le_mois)+"'\
                    and to_char(date,'YYYY')='"+str(annee)+"'\
                    and pres.category_id=(select id from hr_employee_category where name='TA')\
                    group by emp.complete_name\
                    order by emp.complete_name",(tuple(ids),))
         
        for res1 in cr.fetchall():
            h_pres+='+B'+str(x)
            h_for+='+C'+str(x)
            h_brief+='+D'+str(x)
            h_panne+='+E'+str(x)
            h_relance+='+F'+str(x)
            h_fact+='+G'+str(x)
            
            h_pres_tot+='+B'+str(x)
            h_for_tot+='+C'+str(x)
            h_brief_tot+='+D'+str(x)
            h_panne_tot+='+E'+str(x)
            h_relance_tot+='+F'+str(x)
            h_fact_tot+='+G'+str(x)
            
            complete_name=res1[0]
            presence=res1[1]
            finitiale=res1[2]
            brief=res1[3]
            panne=res1[4]
            relance=res1[5]
            feuille_cumul.write('A'+str(x),complete_name,style_nom_ta)
            feuille_cumul.write('B'+str(x),presence,style1_presence_nombre)
            feuille_cumul.write('C'+str(x),finitiale,style1)
            feuille_cumul.write('D'+str(x),brief,style1)
            feuille_cumul.write('E'+str(x),panne,style1_rouge)
            feuille_cumul.write('F'+str(x),relance,style1_chibi)
            feuille_cumul.write_formula('G'+str(x),'=B'+str(x)+'-C'+str(x)+'',style1_vert_nombre)
            x=x+1
        feuille_cumul.write('A'+str(x),'SOUS-TOTAL HEURES TA',style_ss_h_tot_ta)
        feuille_cumul.write_formula('B'+str(x),h_pres[1:],style_ss_h_tot_ta_nombre)
        feuille_cumul.write_formula('C'+str(x),h_for[1:],style_ss_h_tot_ta_nombre)
        feuille_cumul.write_formula('D'+str(x),h_brief[1:],style_ss_h_tot_ta_nombre)
        feuille_cumul.write_formula('E'+str(x),h_panne[1:],style_ss_h_tot_ta_nombre)
        feuille_cumul.write_formula('F'+str(x),h_pnc[1:],style_ss_h_tot_ta_nombre)
        feuille_cumul.write_formula('G'+str(x),h_fact[1:],style_ss_h_tot_ta_nombre)
        x=x+1
        ##################SOUS-TOTAL HEURES Encadrants#########################
        h_pres=h_for=h_brief=h_panne=h_fact=h_relance=heure=''
        cr.execute("select emp.complete_name,sum(case when h_pres is null then 0 else h_pres end),\
                    sum(case when h_formation is null then 0 else h_formation end),\
                    sum(case when h_brief is null then 0 else h_brief end),\
                    sum(case when h_panne is null then 0 else h_panne end),\
                    sum(case when h_relance is null then 0 else h_relance end)\
                    from hr_employee_presence_cdt pres left join hr_employee emp on emp.id=pres.employee_id\
                    where to_char(date,'MM')='"+str(le_mois)+"'\
                    and to_char(date,'YYYY')='"+str(annee)+"'\
                    and pres.category_id in (select id from hr_employee_category where name in ('SUP','Chef de plateau'))\
                    group by emp.complete_name\
                    order by emp.complete_name",(tuple(ids),))
         
        for res1 in cr.fetchall():
            h_pres+='+B'+str(x)
            h_for+='+C'+str(x)
            h_brief+='+D'+str(x)
            h_panne+='+E'+str(x)
            h_relance+='+F'+str(x)
            h_fact+='+G'+str(x)
            
            h_pres_tot+='+B'+str(x)
            h_for_tot+='+C'+str(x)
            h_brief_tot+='+D'+str(x)
            h_panne_tot+='+E'+str(x)
            h_relance_tot+='+F'+str(x)
            h_fact_tot+='+G'+str(x)
            
            complete_name=res1[0]
            presence=res1[1]
            finitiale=res1[2]
            brief=res1[3]
            panne=res1[4]
            relance=res1[5]
            feuille_cumul.write('A'+str(x),complete_name,style_nom_sup)
            feuille_cumul.write('B'+str(x),presence,style1_presence_nombre)
            feuille_cumul.write('C'+str(x),finitiale,style1)
            feuille_cumul.write('D'+str(x),brief,style1)
            feuille_cumul.write('E'+str(x),panne,style1_rouge)
            feuille_cumul.write('F'+str(x),relance,style1_chibi)
            feuille_cumul.write_formula('G'+str(x),'=B'+str(x)+'-C'+str(x)+'-D'+str(x)+'-E'+str(x)+'-F'+str(x)+'',style1_vert_nombre)
            x=x+1
        feuille_cumul.write('A'+str(x),'SOUS-TOTAL HEURES Encadrants',style_ss_h_tot_enc)
        feuille_cumul.write_formula('B'+str(x),h_pres[1:],style_ss_h_tot_enc_nombre)
        feuille_cumul.write_formula('C'+str(x),h_for[1:],style_ss_h_tot_enc_nombre)
        feuille_cumul.write_formula('D'+str(x),h_brief[1:],style_ss_h_tot_enc_nombre)
        feuille_cumul.write_formula('E'+str(x),h_panne[1:],style_ss_h_tot_enc_nombre)
        feuille_cumul.write_formula('F'+str(x),h_pnc[1:],style_ss_h_tot_enc_nombre)
        feuille_cumul.write_formula('G'+str(x),h_fact[1:],style_ss_h_tot_enc_nombre)
        x=x+1
        
        x=x+1
        ##############TOTAL GENERAL############################
        feuille_cumul.write('A'+str(x),'TOTAL GENERAL HEURES  ',style_tot_gen)
        feuille_cumul.write_formula('B'+str(x),h_pres_tot[1:],style_tot_gen_nombre)
        feuille_cumul.write_formula('C'+str(x),h_for_tot[1:],style_tot_gen_nombre)
        feuille_cumul.write_formula('D'+str(x),h_brief_tot[1:],style_tot_gen_nombre)
        feuille_cumul.write_formula('E'+str(x),h_panne_tot[1:],style_tot_gen_nombre)
        feuille_cumul.write_formula('F'+str(x),h_relance_tot[1:],style_tot_gen_nombre)
        feuille_cumul.write_formula('G'+str(x),h_fact_tot[1:],style_tot_gen_nombre)
        
        #################################################################################################################################
        feuille_total=workbook.add_worksheet('Total')
        feuille_total.set_zoom(85)
        feuille_total.set_column('A:A', 42.14)
        feuille_total.set_row(0, 120)
        x=2
        cr.execute("select emp.complete_name,case when emp.state='arret' or emp.state like 'abandon' then 1 else 0 end\
                        from hr_employee_presence_cdt pres left join hr_employee emp on emp.id=pres.employee_id\
                        where to_char(date,'MM')='"+str(le_mois)+"'\
                        and to_char(date,'YYYY')='"+str(annee)+"'\
                        and pres.category_id=(select id from hr_employee_category where name='TA')\
                        group by emp.complete_name,emp.state\
                        order by emp.complete_name",(tuple(ids),))
        for res in cr.fetchall():
            if res[1]==1:
                feuille_total.write('A'+str(x),res[0],style_nom_red)
            else :
                feuille_total.write('A'+str(x),res[0],style_nom_white)
            x=x+1
        cr.execute("select emp.complete_name,case when emp.state='arret' or emp.state like 'abandon' then 1 else 0 end\
                        from hr_employee_presence_cdt pres left join hr_employee emp on emp.id=pres.employee_id\
                        where to_char(date,'MM')='"+str(le_mois)+"'\
                        and to_char(date,'YYYY')='"+str(annee)+"'\
                        and pres.category_id in (select id from hr_employee_category where name in ('SUP','Chef de plateau'))\
                        group by emp.complete_name,emp.state\
                        order by emp.complete_name",(tuple(ids),))
        for res in cr.fetchall():
            if res[1]==1:
                feuille_total.write('A'+str(x),res[0],style_nom_red)
            else :
                feuille_total.write('A'+str(x),res[0],style_nom_white)
            x=x+1
        #feuille_total.write('A'+str(x),'TOTAL',style_nom_white)
        date_=str(annee)+'-'+str(le_mois)+'-01'
        nbr_jour=0
        cr.execute("select \
                    to_char(date(date_trunc('month', date('"+date_+"')))+generate_series\
                    (\
                    0,\
                    (select cast(date_part('day',date_now + '1 month - 1 day'::interval)-1 AS integer)\
                    FROM ( select date(date_trunc('month', date('"+date_+"'))) as date_now) months \
                    )\
                    ), 'D')",(tuple(ids),))
        for res in cr.fetchall():
            nbr_jour+=1
        nbr_contact=0
        cr.execute("select emp.complete_name,case when emp.state='arret' or emp.state like 'abandon' then 1 else 0 end\
                        from hr_employee_presence_cdt pres left join hr_employee emp on emp.id=pres.employee_id\
                        where to_char(date,'MM')='"+str(le_mois)+"'\
                        and to_char(date,'YYYY')='"+str(annee)+"'\
                        and pres.category_id in (select id from hr_employee_category where name in ('TA','SUP','Chef de plateau'))\
                        group by emp.complete_name,emp.state\
                        order by emp.complete_name",(tuple(ids),))
        for res in cr.fetchall():
            nbr_contact+=1
        y=1
        
        cr.execute("select\
                                to_char(date(date_trunc('month', date('"+date_+"')))+generate_series\
                                (\
                                0,\
                                (select cast(date_part('day',date_now + '1 month - 1 day'::interval)-1 AS integer)\
                                FROM ( select date(date_trunc('month', date('"+date_+"'))) as date_now) months\
                                )\
                                ), 'YYYY-MM-DD'),\
                                to_char(date(date_trunc('month', date('"+date_+"')))+generate_series\
                                (\
                                0,\
                                (select cast(date_part('day',date_now + '1 month - 1 day'::interval)-1 AS integer)\
                                FROM ( select date(date_trunc('month', date('"+date_+"'))) as date_now) months \
                                )\
                                ), 'DD'),\
                                to_char(date(date_trunc('month', date('2017-02-01')))+generate_series\
                                (\
                                0,\
                                (select cast(date_part('day',date_now + '1 month - 1 day'::interval)-1 AS integer)\
                                FROM ( select date(date_trunc('month', date('"+date_+"'))) as date_now) months \
                                )\
                                ), 'MM'),\
                                to_char(date(date_trunc('month', date('"+date_+"')))+generate_series\
                                (\
                                0,\
                                (select cast(date_part('day',date_now + '1 month - 1 day'::interval)-1 AS integer)\
                                FROM ( select date(date_trunc('month', date('"+date_+"'))) as date_now) months \
                                )\
                                ), 'D')",(tuple(ids),))
        for res in cr.fetchall():
            x=1
            feuille_total.set_column(self.get_nom_colonne(y)+':'+self.get_nom_colonne(y), 5)
            feuille_total.write(self.get_nom_colonne(y)+str(x),self.get_nom_jour(res[3])+' '+str(res[1])+' '+self.get_nom_mois(res[2])+' '+str(annee),style1_rotation)
            x=x+1
            cr.execute("select distinct emp.complete_name,t.h_pres,case when emp.state='arret' or emp.state like 'abandon' then 1 else 0 end\
                        from hr_employee_presence_cdt pres \
                        left join hr_employee emp on emp.id=pres.employee_id\
                        left join (select employee_id,sum(case when h_pres is null then 0 else h_pres end) as h_pres\
                        from hr_employee_presence_cdt pres\
                        where date='"+res[0]+"'\
                        and pres.category_id=(select id from hr_employee_category where name='TA')\
                        group by employee_id) as t on emp.id=t.employee_id\
                        where to_char(date,'MM')='"+str(le_mois)+"'\
                        and to_char(date,'YYYY')='"+str(annee)+"'\
                        and pres.category_id=(select id from hr_employee_category where name='TA')\
                        order by emp.complete_name",(tuple(ids),))
            for res1 in cr.fetchall():
                if res1[2]==1:
                    feuille_total.write(self.get_nom_colonne(y)+str(x), res1[1],style1_red)
                elif res1[1] is None:
                    feuille_total.write(self.get_nom_colonne(y)+str(x), res1[1],style1_black)
                else:
                    feuille_total.write(self.get_nom_colonne(y)+str(x), res1[1],style1)
                x=x+1
                
            cr.execute("select distinct emp.complete_name,t.h_pres,case when emp.state='arret' or emp.state like 'abandon' then 1 else 0 end\
                        from hr_employee_presence_cdt pres \
                        left join hr_employee emp on emp.id=pres.employee_id\
                        left join (select employee_id,sum(case when h_pres is null then 0 else h_pres end) as h_pres\
                        from hr_employee_presence_cdt pres\
                        where date='"+res[0]+"'\
                        and pres.category_id in (select id from hr_employee_category where name in ('SUP','Chef de plateau'))\
                        group by employee_id) as t on emp.id=t.employee_id\
                        where to_char(date,'MM')='"+str(le_mois)+"'\
                        and to_char(date,'YYYY')='"+str(annee)+"'\
                        and pres.category_id in (select id from hr_employee_category where name in ('SUP','Chef de plateau'))\
                        order by emp.complete_name",(tuple(ids),))
            for res1 in cr.fetchall():
                if res1[2]==1:
                    feuille_total.write(self.get_nom_colonne(y)+str(x), res1[1],style1_red)
                elif res1[1] is None:
                    feuille_total.write(self.get_nom_colonne(y)+str(x), res1[1],style1_black)
                else:
                    feuille_total.write(self.get_nom_colonne(y)+str(x), res1[1],style1)
                x=x+1
            y=y+1
        x=1
        feuille_total.write(self.get_nom_colonne(y)+str(x), "Nombre total d'heures",style)
        x=x+1
        cr.execute("select emp.complete_name\
                        from hr_employee_presence_cdt pres left join hr_employee emp on emp.id=pres.employee_id\
                        where to_char(date,'MM')='"+str(le_mois)+"'\
                        and to_char(date,'YYYY')='"+str(annee)+"'\
                        and pres.category_id = (select id from hr_employee_category where name in ('TA'))\
                        group by emp.complete_name\
                        order by emp.complete_name",(tuple(ids),))
        for res in cr.fetchall():
            feuille_total.write_formula(self.get_nom_colonne(y)+str(x), "SUM(B"+str(x)+":"+self.get_nom_colonne(nbr_jour)+str(x)+")",style_nom_ta)
            x=x+1
        cr.execute("select emp.complete_name\
                        from hr_employee_presence_cdt pres left join hr_employee emp on emp.id=pres.employee_id\
                        where to_char(date,'MM')='"+str(le_mois)+"'\
                        and to_char(date,'YYYY')='"+str(annee)+"'\
                        and pres.category_id in (select id from hr_employee_category where name in ('SUP','Chef de plateau'))\
                        group by emp.complete_name\
                        order by emp.complete_name",(tuple(ids),))
        for res in cr.fetchall():
            feuille_total.write_formula(self.get_nom_colonne(y)+str(x), "SUM(B"+str(x)+":"+self.get_nom_colonne(nbr_jour)+str(x)+")",style_nom_sup)
            x=x+1
        workbook.close()
        return fichier
reporting()
class planning_cdt(osv.osv):
    _name = "planning.cdt"
    _description = "Planning cdt"
    
    _columns = {
       }
                   

    _defaults = {
    }
planning_cdt()