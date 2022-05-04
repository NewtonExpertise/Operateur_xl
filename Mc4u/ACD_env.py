try:
    from newton_env.mariadbagent import Mariadb
except :
    from Mc4u.mariadbagent import Mariadb
from datetime import datetime
from calendar import monthrange


def push_B(func):
    def decorateur(*args, **kwargs):
        with Mariadb("baku") as Mdb:
            data = Mdb.insert(func(*args, **kwargs))
        return data
    return decorateur

def execute_B(func):
    def decorateur(*args, **kwargs):
        with Mariadb("baku") as Mdb:
            data = Mdb.execute(func(*args, **kwargs))
        return data
    return decorateur

def execute_B_with_header(func):
    def decorateur(*args, **kwargs):
        with Mariadb("baku") as Mdb:
            data = Mdb.execute_with_header(func(*args, **kwargs))
        return data
    return decorateur

def push_R(func):
    def decorateur(*args, **kwargs):
        with Mariadb("rome") as Mdb:
            data = Mdb.insert(func(*args, **kwargs))
        return data
    return decorateur

def execute_R(func):
    def decorateur(*args, **kwargs):
        with Mariadb("rome") as Mdb:
            data = Mdb.execute(func(*args, **kwargs))
        return data
    return decorateur

def get_one_result(func):
    def deco(*args, **kwargs):
        x = func(*args, **kwargs)
        if x:
            return x[0][0]
        else:
            return None
    return deco

def get_list_result(func):
    def deco(*args, **kwargs):
        x = func(*args, **kwargs)
        if x:
            if len(x) == 1:
                data = []
                data.append(list(x[0]))
            else:
                if len(x[0]) ==1:
                    data = [y[0] for y in x]
                else:
                    data = [list(y)for y in x]
            return data
        else:
            return None
    return deco

def check_if_exist(func):
    def deco(*args, **kwargs):
        if func(*args, **kwargs):
            return True
        else:
            return False
    return deco



@get_one_result
@execute_R
def get_raison_social(code_client):
    
    sql =f"""select adr_nom FROM expert.adresse where ADR_CODE = '{code_client}'"""
    return sql

@get_list_result
@execute_R
def get_list_client():
    sql ="""select ADR_CODE, ADR_NOM
    from expert.adresse
    where soc_code = 'NEWTONEXP'
    AND GENRE_CODE like 'CLI%'"""
    return sql


def get_dict_list_client():
    _dict = {}
    data = get_list_client()
    for x, y in data:
        _dict.update({f"{y} ({x})":x})
    return _dict

@get_list_result
@execute_B
def get_dataBase_acd_cpta():
    sql="SHOW DATABASES"
    return sql


@get_one_result
@execute_R
def get_id_dossier(code_client):
    sql =f"""select ADR_ID FROM expert.adresse where ADR_CODE = '{code_client}'"""
    return sql

@get_list_result
@execute_R
def get_dates_exercice(code_dossier):
    ARD_id = get_id_dossier(code_dossier)
    sql =f"""select EXO_DATE_DEB, EXO_DATE_FIN FROM expert.exercice where ADR_ID = '{ARD_id}'"""
    return sql


@get_one_result
@execute_R
def get_id_dossier(code_client):
    sql =f"""select ADR_ID FROM expert.adresse where ADR_CODE = '{code_client}'"""
    return sql

@get_list_result
@execute_R
def get_dates_exercice_courant(ADR_ID):
    sql =f"""select EXO_DATE_DEB, EXO_DATE_FIN FROM expert.exercice where ADR_ID = '{ADR_ID}' and exo_b_courant = 1"""
    return sql

def get_exercice_en_cours(code_client):
    _id_dossier = get_id_dossier(code_client)
    debut , fin= get_dates_exercice_courant(_id_dossier)
    debut = datetime.strptime(debut, '%Y%m%d')
    fin = datetime.strptime(fin, '%Y%m%d')
    return debut, fin

@get_one_result
@execute_R
def get_annee_exercice_courant(ADR_ID):
    sql =f"""select EXO_CODE FROM expert.exercice where ADR_ID = '{ADR_ID}' and exo_b_courant = 1"""
    return sql


@execute_R
def get_exercice_ouvert(code_client):
    _id_dossier = get_id_dossier(code_client)
    annee = get_annee_exercice_courant(_id_dossier)
    sql = f"""select EXO_DATE_DEB, EXO_DATE_FIN from expert.exercice where ADR_ID = '{_id_dossier}' and EXO_code >= '{annee}' order by exo_date_fin desc """
    return sql

@execute_R
def get_all_exercice(code_client):
    _id_dossier = get_id_dossier(code_client)
    sql = f"""select EXO_DATE_DEB, EXO_DATE_FIN from expert.exercice where ADR_ID = '{_id_dossier}' order by exo_date_fin desc """
    return sql

# @execute_B
@get_list_result
@execute_B
def get_ana_exercice(code_client, debut, fin, raison_social = None):
    debut = debut
    fin = fin
    if raison_social:
        sql = f"""select * from (select e.JNL_CODE as Journal,
            DATE(CONCAT_WS('-', e.ECR_ANNEE, e.ECR_MOIS, le.LE_JOUR)) as Date,
            le.CPT_CODE as Compte,
            le.LE_LIB as Libelle,
            A.ANA_DEB_ORG  as Debit_ana,
            A.ANA_CRE_ORG as Credit_ana,
            A.ANA_DEB_ORG - a.ANA_CRE_ORG as Solde_ana,
            le.LE_PIECE as Pièce,
            a.ANA_N1 as Centre,
            le.LE_DEB_ORG as Debit,
            le.LE_CRE_ORG as Credit,
            le.LE_DEB_ORG - le.LE_CRE_ORG as Solde,
            ABS(A.ANA_DEB_ORG - a.ANA_CRE_ORG) - ABS(le.LE_DEB_ORG - le.LE_CRE_ORG) as compar_gene_ana,
            '{raison_social}'
        from compta_{code_client}.ecriture e , compta_{code_client}.ligne_ecriture le, compta_{code_client}.analytique a
        where e.ECR_CODE = le.ECR_CODE
        and le.LE_CODE = a.LE_CODE
    union 
    select he.JNL_CODE as Journal,
            DATE(CONCAT_WS('-', he.HE_ANNEE, he.HE_MOIS, hle.HLE_JOUR)) as Date,
            hle.CPT_CODE as Compte,
            hle.HLE_LIB as Libelle,
            ha.HANA_DEB_ORG  as Debit_ana,
            ha.HANA_CRE_ORG as Credit_ana,
            ha.HANA_DEB_ORG - ha.HANA_CRE_ORG as Solde_ana,
            hle.HLE_PIECE as Pièce,
            ha.HANA_N1 as Centre,
            hle.HLE_DEB_ORG as Debit,
            hle.HLE_CRE_ORG as Credit,
            hle.HLE_DEB_ORG - hle.HLE_CRE_ORG as Solde,
            ABS(ha.HANA_DEB_ORG - ha.HANA_CRE_ORG) - ABS(hle.HLE_DEB_ORG - hle.HLE_CRE_ORG) as compar_gene_ana,
            '{raison_social}'
        from compta_{code_client}.histo_analytique ha , compta_{code_client}.histo_ecriture he , compta_{code_client}.histo_ligne_ecriture hle 
        where he.HE_CODE = hle.HE_CODE
        and hle.HLE_CODE = ha.HLE_CODE) as x 
        where x.Date between '{debut}' and '{fin}'
        and (x.Compte like '6%' or x.Compte like '7%')"""
    else:
        sql = f"""select * from (select e.JNL_CODE as Journal,
            DATE(CONCAT_WS('-', e.ECR_ANNEE, e.ECR_MOIS, le.LE_JOUR)) as Date,
            le.CPT_CODE as Compte,
            le.LE_LIB as Libelle,
            A.ANA_DEB_ORG  as Debit_ana,
            A.ANA_CRE_ORG as Credit_ana,
            A.ANA_DEB_ORG - a.ANA_CRE_ORG as Solde_ana,
            le.LE_PIECE as Pièce,
            a.ANA_N1 as Centre,
            le.LE_DEB_ORG as Debit,
            le.LE_CRE_ORG as Credit,
            le.LE_DEB_ORG - le.LE_CRE_ORG as Solde,
            ABS(A.ANA_DEB_ORG - a.ANA_CRE_ORG) - ABS(le.LE_DEB_ORG - le.LE_CRE_ORG) as compar_gene_ana
        from compta_{code_client}.ecriture e ,compta_{code_client}.ligne_ecriture le, compta_{code_client}.analytique a
        where e.ECR_CODE = le.ECR_CODE
        and le.LE_CODE = a.LE_CODE
    UNION ALL
    select he.JNL_CODE as Journal,
            DATE(CONCAT_WS('-', he.HE_ANNEE, he.HE_MOIS, hle.HLE_JOUR)) as Date,
            hle.CPT_CODE as Compte,
            hle.HLE_LIB as Libelle,
            ha.HANA_DEB_ORG  as Debit_ana,
            ha.HANA_CRE_ORG as Credit_ana,
            ha.HANA_DEB_ORG - ha.HANA_CRE_ORG as Solde_ana,
            hle.HLE_PIECE as Pièce,
            ha.HANA_N1 as Centre,
            hle.HLE_DEB_ORG as Debit,
            hle.HLE_CRE_ORG as Credit,
            hle.HLE_DEB_ORG - hle.HLE_CRE_ORG as Solde,
            ABS(ha.HANA_DEB_ORG - ha.HANA_CRE_ORG) - ABS(hle.HLE_DEB_ORG - hle.HLE_CRE_ORG) as compar_gene_ana
        from compta_{code_client}.histo_analytique ha , compta_{code_client}.histo_ecriture he , compta_{code_client}.histo_ligne_ecriture hle 
        where he.HE_CODE = hle.HE_CODE
        and hle.HLE_CODE = ha.HLE_CODE) as x 
        where x.Date between '{debut}' and '{fin}'
        and (x.Compte like '6%' or x.Compte like '7%')"""

    return sql



@get_list_result
@execute_B
def get_balance_ana_exercice(code_client, debut, fin):
    debut = debut
    cloture = last_day_of_month(fin.month, fin.year)
    sql = f"""
select CODE_ANA_cumul, mensuel, cumul from (
select CODE_ANA_cumul, SUM(DEBIT_cumul - CREDIT_cumul) as cumul
from ( 
		select a.ANA_N1 AS CODE_ANA_cumul, a.ANA_DEB_ORG as DEBIT_cumul, a.ANA_CRE_ORG AS CREDIT_cumul, STR_TO_DATE(CONCAT(le.LE_JOUR,'/',e.ECR_MOIS,'/',e.ECR_ANNEE),'%d/%m/%Y') as Date_Ecriture
		from compta_{code_client}.analytique a
		left join compta_{code_client}.ligne_ecriture le on le.LE_CODE = a.LE_CODE 
		left join compta_{code_client}.ecriture e on e.ECR_CODE = le.ecr_CODE
		where STR_TO_DATE(CONCAT(le.LE_JOUR,'/',e.ECR_MOIS,'/',e.ECR_ANNEE),'%d/%m/%Y') between '{debut}' and '{cloture}'
		and (le.CPT_CODE like '6%' or le.CPT_CODE like '7%')
		union all 
		select ha.hANA_N1 as CODE_ANA_cumul, ha.hANA_DEB_ORG as DEBIT_cumul, ha.hANA_CRE_ORG AS CREDIT_cumul, STR_TO_DATE(CONCAT(hle.hLE_JOUR,'/',he.hE_MOIS,'/',he.hE_ANNEE),'%d/%m/%Y') as Date_Ecriture
		from compta_{code_client}.histo_analytique ha
		left join compta_{code_client}.histo_ligne_ecriture hle on hle.HLE_CODE = ha.HLE_CODE 
		left join compta_{code_client}.histo_ecriture he on he.HE_CODE = hle.HE_CODE
		where STR_TO_DATE(CONCAT(hle.hLE_JOUR,'/',he.hE_MOIS,'/',he.hE_ANNEE),'%d/%m/%Y') between '{debut}' and '{cloture}'
		and (hle.CPT_CODE like '6%' or hle.CPT_CODE like '7%')) as subquery_cumul
		group by subquery_cumul.CODE_ANA_cumul) as rt_cumul
left join ( select CODE_ANA_mensuel, SUM(DEBIT_mensuel-CREDIT_mensuel) as mensuel
				from(
					select a.ANA_N1 AS CODE_ANA_mensuel, a.ANA_DEB_ORG as DEBIT_mensuel, a.ANA_CRE_ORG AS CREDIT_mensuel, STR_TO_DATE(CONCAT(le.LE_JOUR,'/',e.ECR_MOIS,'/',e.ECR_ANNEE),'%d/%m/%Y') as Date_Ecriture
					from compta_{code_client}.analytique a
					left join compta_{code_client}.ligne_ecriture le on le.LE_CODE = a.LE_CODE 
					left join compta_{code_client}.ecriture e on e.ECR_CODE = le.ecr_CODE
					where STR_TO_DATE(CONCAT(le.LE_JOUR,'/',e.ECR_MOIS,'/',e.ECR_ANNEE),'%d/%m/%Y') between '{fin}' and '{cloture}'
					and (le.CPT_CODE like '6%' or le.CPT_CODE like '7%')
					union all 
					select ha.hANA_N1 as CODE_ANA_mensuel, ha.hANA_DEB_ORG as DEBIT_mensuel, ha.hANA_CRE_ORG AS CREDIT_mensuel, STR_TO_DATE(CONCAT(hle.hLE_JOUR,'/',he.hE_MOIS,'/',he.hE_ANNEE),'%d/%m/%Y') as Date_Ecriture
					from compta_{code_client}.histo_analytique ha
					left join compta_{code_client}.histo_ligne_ecriture hle on hle.HLE_CODE = ha.HLE_CODE 
					left join compta_{code_client}.histo_ecriture he on he.HE_CODE = hle.HE_CODE
					where STR_TO_DATE(CONCAT(hle.hLE_JOUR,'/',he.hE_MOIS,'/',he.hE_ANNEE),'%d/%m/%Y') between '{fin}' and '{cloture}'
					and (hle.CPT_CODE like '6%' or hle.CPT_CODE like '7%')) as ana_mensuel
					group by ana_mensuel.CODE_ANA_mensuel
			) mensu on mensu.CODE_ANA_mensuel = rt_cumul.CODE_ANA_cumul
      """

    return sql

@get_list_result
@execute_B
def get_journaux_dossier(code_client):
    sql = f"""select * 
    from    (select e.JNL_CODE as Journal
            from {code_client}.ecriture e
            union
            select he.JNL_CODE as Journal
            from {code_client}.histo_ecriture he) as e
    group by e.Journal"""
    return sql

@get_list_result
@execute_R
def get_list_group_mcdo(code_client):
    sql = f"""select Concat('compta_',a.adr_code), a.adr_nom from expert.adresse a, expert.groupe g 
            where g.GRP_CODE = a.GRP_CODE 
            and g.GRP_CODE = (select a2.GRP_CODE from expert.adresse a2 where a2.ADR_CODE = '{code_client}')
            and a.NAF_CODE = '5610C'"""
    return sql

@get_one_result
@execute_B
def get_holding_result(codeclient, debut, fin ):
    fin = last_day_of_month(fin.month, fin.year)
    sql = f"""select sum(c)-SUM(d)
from(
select (le.LE_DEB_ORG - le.LE_CRE_ORG) as d, 0 as c from compta_{codeclient}.ecriture e, compta_{codeclient}.ligne_ecriture le 
where le.ECR_CODE = e.ECR_CODE 
and le.CPT_CODE like '6%'
and DATE(CONCAT_WS('-', e.ECR_ANNEE, e.ECR_MOIS, le.LE_JOUR)) between '{debut}' and '{fin}'
union all 
select 0 as d, (le.LE_CRE_ORG - le.LE_DEB_ORG) as c from compta_{codeclient}.ecriture e, compta_{codeclient}.ligne_ecriture le 
where le.ECR_CODE = e.ECR_CODE 
and le.CPT_CODE like '7%'
and DATE(CONCAT_WS('-', e.ECR_ANNEE, e.ECR_MOIS, le.LE_JOUR)) between '{debut}' and '{fin}'
union all
select (hle2.HLE_DEB_ORG - hle2.HLE_CRE_ORG) as d, 0 as c from compta_{codeclient}.histo_ecriture he2 , compta_{codeclient}.histo_ligne_ecriture hle2 
where he2.HE_CODE = hle2.HE_CODE
and hle2.CPT_CODE like '6%'
and DATE(CONCAT_WS('-', he2.HE_ANNEE , he2.HE_MOIS , hle2.HLE_JOUR)) between '{debut}' and '{fin}'
union all 
select 0 as d, (hle2.HLE_CRE_ORG - hle2.HLE_DEB_ORG) as c from compta_{codeclient}.histo_ecriture he2, compta_{codeclient}.histo_ligne_ecriture hle2 
where he2.HE_CODE = hle2.HE_CODE
and hle2.CPT_CODE like '7%'
and DATE(CONCAT_WS('-', he2.HE_ANNEE , he2.HE_MOIS , hle2.HLE_JOUR)) between '{debut}' and '{fin}') as x"""

    return sql 


@get_one_result
@execute_R
def get_group_mcdo(code_client):
    sql = f"""select concat(g.GRP_NOM,' ', g.GRP_PRENOM) as nom_group
from expert.adresse a, expert.groupe g 
where g.GRP_CODE = a.GRP_CODE 
and a.ADR_CODE = '{code_client}'"""
    return sql

def last_day_of_month(month, year):
    date_value = datetime(year, month, 1)
    return date_value.replace(day = monthrange(date_value.year, date_value.month)[1])

def get_month_period(start, end):
    total_months = lambda dt: dt.month + 12 * dt.year
    mlist = []
    for tot_m in range(total_months(start)-1, total_months(end)):
        y, m = divmod(tot_m, 12)
        mlist.append(datetime(y, m+1, 1).strftime("%B %Y"))
    mlist = [date.capitalize() for date in mlist]

    return mlist
