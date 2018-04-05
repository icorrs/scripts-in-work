#export detail_wbs,which give to project-part;
import sqlalchemy
import pymysql
import pandas as pd

from cost_analysis import translate_title

date=input('enter date tobe export:')
engine=sqlalchemy.create_engine('mysql+pymysql://root:nakamura7@localhost:3306/cost_analysis?charset=utf8')


def get_detail_wbs(engine,date):
    'main attrbute'
    sql_query='''
                select detail_wbs.detail_wbs_content,
                       detail_wbs.detail_wbs_code,
                       detail_wbs.detail_wbs_beginning_mileage,
                       detail_wbs.detail_wbs_ending_mileage,
                       detail_wbs.detail_wbs_quantity,
                       %s_quantity.%s_quantity,
                       detail_wbs_code_vs_sub_contractor_code.sub_contractor_short_name
                from detail_wbs,%s_quantity,detail_wbs_code_vs_sub_contractor_code
                where detail_wbs.detail_wbs_code=%s_quantity.detail_wbs_code and 
                      detail_wbs_code_vs_sub_contractor_code.detail_wbs_code=detail_wbs.detail_wbs_code
                '''%(date,date,date,date)
    frame1=pd.read_sql_query(sql_query,engine)
    frame1=translate_title(frame1)
    frame1.to_csv(r'c:\python tem\detail_wbs_export.csv',encoding='utf-8-sig')


if __name__=='__main__':
    get_detail_wbs(engine,date)
