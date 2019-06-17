# import cast_upgrade_1_5_21  #@UnusedImport
import os
import time
import xlsxwriter
from cast.application import ApplicationLevelExtension
import json
import requests
import xml.etree.ElementTree as ET
from cast.application import create_postgres_engine
# from numpy.distutils.fcompiler import None
from decimal import Decimal
import cast.analysers.log as PRINT


class Report(ApplicationLevelExtension):
    
    def __init__(self):
        
        pass
    
    def start_application(self, application):
#         kb = application.get_knowledge_base()
#         kb.execute_query("""delete from SYS_SITE_OPTIONS where OPTION_NAME = 'GRAPH_SAVE_LARGEST_SCC_GROUP';""")
#         kb.execute_query("""Insert into SYS_SITE_OPTIONS(OPTION_NAME, OPTION_VALUE) values ('GRAPH_SAVE_LARGEST_SCC_GROUP', '10');""")
        pass
    
    def end_application(self, application):
        pass
   
    def after_snapshot(self, application):

        from cast.application import publish_report  # @UnresolvedImport
        
        # generate a path in LISA my_report<timestamp>.xlxs 
        report_path = os.path.join(self.get_plugin().intermediate, time.strftime("Rescan Audit Report%Y%m%d_%H%M%S.xlsx"))

        workbook = xlsxwriter.Workbook(report_path)
                          
        # kb = application.get_knowledge_base()
        kb = application.get_application_configuration().get_analysis_service()        
        central = application.get_central() 
        mngt = application.get_managment_base()
       
        # rptpth = 'E:\CASTMS\DMTDelivery836\data\{068916b1-9f32-453d-a14e-1f960dccde0c}\{6a05cfb9-868e-4b20-ae15-4f39b2af6705}\DMTDeliveryReport.xml'
       
        worksheet = workbook.add_worksheet('Rescan Audit')
        bold = workbook.add_format({'bold': True , 'align': 'left'})
        bold.set_bg_color('#9BC2E6')
        hdngclr = workbook.add_format({'bold': True, 'align': 'left'})
        hdngclr.set_bg_color('#FFC000')
        ylclr = workbook.add_format({'bold': True, 'align': 'left'})
        ylclr.set_bg_color('#8EA9DB')
        dtfrmt = workbook.add_format({'num_format': 'mm/dd/yy', 'align': 'left'})
        adtval_algn = workbook.add_format({'align': 'left'})
        adtval_algn1 = workbook.add_format({'align': 'left', 'color': 'red'})  
              
        # clr = workbook.add_format({'bg_color': 'cyan'})    
        worksheet.set_column(0, 5, 40)
            
        dlmSts = ''        
        varPercent = ''
        crntAEFP = ''
        crntAEFT = ''
        prvAEFP = ''
        prvAEFT = ''
        crntfp = ''
        prcnt = ''   
        consval_metid = ''
        consval_metName = ''
        consval_prev = ''
        consval_crnt = ''
        consval_prcnt = ''   
        met_add = ''
        met_lost = ''
        scc_num = ''
        object_id = ''
        artfcts = ''  
        dltd = ''
        aded = ''
        comn = ''
        obj_lang = ''
        lst_cnt = ''
        cnt = ''
        techName = ''
        snapshtName = ''
        worksheet.write(0, 0, 'Application Name', bold)
        
        for line6 in mngt.execute_query("""Select object_name from cms_portf_application;"""): 
            if line6[0]: 
                worksheet.write(0, 1, line6[0])
        worksheet.write(1, 0, 'AUDIT CHECKS', hdngclr)
        worksheet.write(1, 1, 'Check Value', hdngclr)
        worksheet.write(1, 2, 'Check Status', hdngclr)
        worksheet.write(0, 2, 'Snapshot Name', bold)
        worksheet.write(0, 4, 'Snapshot Date', bold)
        for line6 in central.execute_query("""select snapshot_name as "Snapshot Name" from dss_snapshots where snapshot_id in (select max(snapshot_id) from dss_snapshots);"""): 
            if line6[0] is not None: 
                worksheet.write(0, 3, line6[0])
        for line6 in central.execute_query("""select snapshot_date as "Snapshot Date" from dss_snapshots where snapshot_id in (select max(snapshot_id) from dss_snapshots);"""): 
            if line6[0] is not None: 
                worksheet.write(0, 5, line6[0], dtfrmt)        
        
        worksheet.write(1, 3, 'AUDIT INFO', hdngclr)
        worksheet.write(1, 4, 'AUDIT VALUE', hdngclr)
        worksheet.write(2, 3, 'Previous AEFP;AETP', bold)
        worksheet.write(3, 3, 'Current FP', bold)
        worksheet.write(4, 3, 'Artifact coverage Current version', bold)
        worksheet.write(5, 3, 'Arififact coverge Previous version', bold)
        worksheet.write(6, 3, 'Empty Transaction Current', bold)
        worksheet.write(7, 3, 'Empty Transaction Previous', bold)
        worksheet.write(8, 3, 'Current AEFP', bold)
        worksheet.write(9, 3, 'Current AETP', bold)
        
#------------Previous AEFP;AETP------------------------------------        
        worksheet.write(2, 4, 'NA')
        for line6 in central.execute_query("""select c3.metric_num_value as "Previous AEFP", c4.metric_num_value as "Previous AETP"  from dss_metric_results c1, dss_metric_results c2,dss_metric_results c3, dss_metric_results c4
where c1.metric_id  = 10430 and c2.metric_id = 10440 and c3.metric_id  = 10430 and c4.metric_id = 10440 
and c1.object_id in (select object_id from dss_objects where object_type_id =-102) and c2.object_id in (select object_id from dss_objects where object_type_id =-102) 
and c1.snapshot_id = (select max(snapshot_id) from dss_snapshots) and c2.snapshot_id = (select max(snapshot_id) from dss_snapshots) 
and c3.snapshot_id = (select snapshot_id from (select snapshot_id from dss_snapshots order by snapshot_id desc limit 2) as snapid order by snapshot_id limit 1)
and c3.object_id in (select object_id from dss_objects where object_type_id =-102) and c4.object_id in (select object_id from dss_objects where object_type_id =-102) 
and c4.snapshot_id = (select snapshot_id from (select snapshot_id from dss_snapshots order by snapshot_id desc limit 2) as snapid order by snapshot_id limit 1);"""): 
            if line6[0] is not None: 
                prvAEFP = line6[0]
            if line6[1] is not None:
                prvAEFT = line6[1]                
            worksheet.write(2, 4, str(prvAEFP) + ';' + str(prvAEFT))
#------------Current FP------------------------------------     
        worksheet.write(3, 4, 'NA')   
        for line6 in central.execute_query("""select c1.metric_num_value as "Current FP" from dss_metric_results c1 where c1.metric_id  = 10202 
and c1.object_id in (select object_id from dss_objects where object_type_id =-102) and c1.snapshot_id = (select max(snapshot_id) from dss_snapshots);"""):
            if line6[0] is not None: 
                crntfp = line6[0]             
            worksheet.write(3, 4, crntfp, adtval_algn)
#------------Artifact coverage Current version   and ---Arififact coverge Previous version---------------------------------         
        worksheet.write(4, 4, 'NA')
        worksheet.write(5, 4, 'NA')  
        for line6 in central.execute_query("""select distinct c1.percentage as "Artefact Coverage Current Version", c2.percentage as "Artefact Coverage Previous Version"
from aia_console.smelltest_history c1,aia_console.smelltest_history c2    where c1.currentversion = (select snapshot_name from dss_snapshots where snapshot_id in (select max(snapshot_id) from dss_snapshots))    
and c2.currentversion = (select snapshot_name from dss_snapshots where snapshot_id in ((select snapshot_id from (select snapshot_id from dss_snapshots order by snapshot_id desc limit 2) as snapid order by 
snapshot_id limit 1))) and c1.checknumber = '1smell' and c2.checknumber = '1smell' and c1.appname = c2.appname;"""):
            if line6[0] is not None: 
                crntAEFP = line6[0]
            if line6[1] is not None:
                crntAEFP = line6[1]                
            worksheet.write(4, 4, str(crntAEFP), adtval_algn)
            worksheet.write(5, 4, str(crntAEFP), adtval_algn)
#------------Empty Transaction Current   and ---Empty Transaction Previous---------------------------------            
        worksheet.write(6, 4, 'NA') 
        worksheet.write(7, 4, 'NA') 
        for line6 in central.execute_query("""select distinct c1.percentage as "Empty Transaction Current", c2.percentage as "Empty Transaction Previous" from aia_console.smelltest_history c1,aia_console.smelltest_history c2    where c1.currentversion = 
(select snapshot_name from dss_snapshots where snapshot_id in (select max(snapshot_id) from dss_snapshots))    
and c2.currentversion = (select snapshot_name from dss_snapshots where snapshot_id in ((select snapshot_id from (select snapshot_id from dss_snapshots order by snapshot_id desc limit 2) as snapid order by 
snapshot_id limit 1))) and c1.checknumber = '2smell' and c2.checknumber = '2smell' and c1.appname = c2.appname;"""):
            if line6[0] is not None: 
                prvAEFP = line6[0]
            if line6[1] is not None:
                prvAEFT = line6[1]                
            worksheet.write(6, 4, str(prvAEFP), adtval_algn)  
            worksheet.write(7, 4, str(prvAEFT), adtval_algn)    

#-------------Current AEFP------------------------------------- 
        worksheet.write(8, 4, 'NA') 
         
        for line3 in central.execute_query("""select c1.metric_num_value as "Current AEFP"
 from 
dss_metric_results c1, dss_metric_results c2,dss_metric_results c3, dss_metric_results c4
where c1.metric_id  = 10430 and c2.metric_id = 10440 
and c3.metric_id  = 10430 and c4.metric_id = 10440 
and c1.object_id in (select object_id from dss_objects where object_type_id =-102) 
and c2.object_id in (select object_id from dss_objects where object_type_id =-102) 
and c1.snapshot_id = (select max(snapshot_id) from dss_snapshots)
and c2.snapshot_id = (select max(snapshot_id) from dss_snapshots) 
and c3.snapshot_id = (select snapshot_id from (select snapshot_id from dss_snapshots order by snapshot_id desc limit 2) as snapid order by snapshot_id limit 1)
and c3.object_id in (select object_id from dss_objects where object_type_id =-102) 
and c4.object_id in (select object_id from dss_objects where object_type_id =-102) 
and c4.snapshot_id = (select snapshot_id from (select snapshot_id from dss_snapshots order by snapshot_id desc limit 2) as snapid order by snapshot_id limit 1);"""):
            if line3[0] is not None:
                crntAEFP = line3[0]                  
            crntAEFP = round(crntAEFP, 2)          
            worksheet.write(8, 4, crntAEFP, adtval_algn)            
                
#-------------Current AETP------------------------------------- 
        worksheet.write(9, 4, 'NA')
        for line3 in central.execute_query("""select c2.metric_num_value as "Current AETP" from 
dss_metric_results c1, dss_metric_results c2,dss_metric_results c3, dss_metric_results c4
where c1.metric_id  = 10430 and c2.metric_id = 10440 
and c3.metric_id  = 10430 and c4.metric_id = 10440 
and c1.object_id in (select object_id from dss_objects where object_type_id =-102) 
and c2.object_id in (select object_id from dss_objects where object_type_id =-102) 
and c1.snapshot_id = (select max(snapshot_id) from dss_snapshots)
and c2.snapshot_id = (select max(snapshot_id) from dss_snapshots) 
and c3.snapshot_id = (select snapshot_id from (select snapshot_id from dss_snapshots order by snapshot_id desc limit 2) as snapid order by snapshot_id limit 1)
and c3.object_id in (select object_id from dss_objects where object_type_id =-102) 
and c4.object_id in (select object_id from dss_objects where object_type_id =-102) 
and c4.snapshot_id = (select snapshot_id from (select snapshot_id from dss_snapshots order by snapshot_id desc limit 2) as snapid order by snapshot_id limit 1);"""):
            if line3[0] is not None:
                crntAEFT = line3[0]               
            crntAEFT = round(crntAEFT, 2)                              
            worksheet.write(9, 4, crntAEFT, adtval_algn)            
            
#----------------------------Unreviewed DLMs Count-------------------                  
        worksheet.write(2, 0, 'Unreviewed DLMs Count', bold)
        worksheet.write(2, 1, 'NA', adtval_algn)
        worksheet.write(2, 1, 'NA', adtval_algn)
        for line1 in kb.execute_query("""SELECT COUNT(prop) FROM acc where prop =1;"""):            
            if line1[0] is not None:
                if line1[0] > 0:
                    dlmSts = line1[0]
                    worksheet.write(2, 1, dlmSts, adtval_algn)
                    worksheet.write(2, 2, "FAILED", adtval_algn1)
                else:
                    dlmSts = line1[0]
                    worksheet.write(2, 1, dlmSts, adtval_algn)      
                    worksheet.write(2, 2, "PASSED", adtval_algn)  
            else:
                worksheet.write(2, 1, 'NA', adtval_algn)
                worksheet.write(2, 1, 'NA', adtval_algn)
            
#----------------------------Dead Code delta for latest 2 snapshot(%)------------------- 
        worksheet.write(3, 0, 'Dead Code delta for latest 2 snapshot(%)', bold)
        worksheet.write(3, 1, 'NA', adtval_algn)
        worksheet.write(3, 2, 'NA', adtval_algn) 
        for line2 in central.execute_query("""select ((SELECT case when Sum(dmr.metric_num_value)::integer <> 0 then round ((c2.metric_num_value/Sum(dmr.metric_num_value))*100,2) else 0 end "Current Dead Code %"
      
FROM   dss_metric_results DMR,
       dss_module_links DML,
       dss_objects DOB,
       dss_metric_results c2
WHERE  dml.object_type_id = 20000
       AND dml.module_id = dmr.object_id
       AND dml.snapshot_id = dmr.snapshot_id
       AND dmr.metric_id = 7832
       AND dmr.metric_value_index = 2
       and c2.object_id in (select object_id from dss_objects where object_type_id =-102) and c2.snapshot_id = dmr.snapshot_id and c2.metric_value_index=1
       and c2.metric_id = 7832
       AND dmr.object_id = DOB.object_id
       AND dmr.snapshot_id = (select max(snapshot_id) from dss_snapshots)
GROUP  BY 
         c2.metric_num_value)
-
(SELECT distinct
     case when Sum(dmr.metric_num_value)::integer <> 0 then round ((c2.metric_num_value/Sum(dmr.metric_num_value))*100,2) else 0 end "Previous Dead Code %"
FROM   dss_metric_results DMR,
       dss_module_links DML,
       dss_objects DOB,
       dss_metric_results c2
WHERE  dml.object_type_id = 20000
       AND dml.module_id = dmr.object_id
       AND dml.snapshot_id = dmr.snapshot_id
       AND dmr.metric_id = 7832
       AND dmr.metric_value_index = 2
       and c2.object_id in (select object_id from dss_objects where object_type_id =-102) and c2.snapshot_id = dmr.snapshot_id and c2.metric_value_index=1
       and c2.metric_id = 7832
       AND dmr.object_id = DOB.object_id
       AND dmr.snapshot_id = (select snapshot_id from (select snapshot_id from dss_snapshots order by snapshot_id desc limit 2) as snapid order by 
                      snapshot_id limit 1)
GROUP  BY 
         c2.metric_num_value)) as "Variation in Dead Code %";"""):
            
            if line2[0] is not None:                
                if line2[0] >= 5:                    
                    # dlmSts = line2[0]
                    dlmSts = round(dlmSts, 2)        
                    worksheet.write(3, 1, str(dlmSts) + '%', adtval_algn)             
                    worksheet.write(3, 2, "FAILED", adtval_algn1)
                else:
                    # dlmSts = line2[0]  
                    dlmSts = round(dlmSts, 2) 
                    worksheet.write(3, 1, str(dlmSts) + '%', adtval_algn)    
                    worksheet.write(3, 2, "PASSED", adtval_algn)
                
        #----------------------------Common Object-------------------    
        worksheet.write(4, 0, 'Common Object(Added & Deleted', bold)
        worksheet.write(4, 1, 'NA', adtval_algn)
        worksheet.write(4, 2, 'NA', adtval_algn)
        for line2 in central.execute_query("""select count(distinct object_full_name) from DSS_OBJECTS o where o.object_full_name in (
SELECT distinct object_full_name FROM 
ADGV_COST_STATUSES cs, DSS_LINKS l, DSS_MODULE_LINKS m, DSS_OBJECTS o, DSS_OBJECT_TYPES ot WHERE m.SNAPSHOT_ID = 
(select max(snapshot_id) from dss_snapshots)
AND cs.object_id = o.object_id AND o.object_type_id = ot.object_type_id AND l.PREVIOUS_OBJECT_ID = m.MODULE_ID AND l.LINK_TYPE_ID = 3 
AND cs.OBJECT_ID = l.NEXT_OBJECT_ID AND cs.SNAPSHOT_ID in (select max(snapshot_id) from dss_snapshots) AND cs.CHANGE_TYPE = 1                    
intersect
SELECT distinct object_full_name
FROM ADGV_COST_STATUSES cs, DSS_LINKS l, DSS_MODULE_LINKS m, DSS_OBJECTS o, DSS_OBJECT_TYPES ot WHERE 
(m.SNAPSHOT_ID in (select max(snapshot_id) from dss_snapshots) or m.SNAPSHOT_ID in (select snapshot_id from (select snapshot_id from dss_snapshots order by snapshot_id desc limit 2) as snapid order by 
                      snapshot_id limit 1))
AND cs.object_id = o.object_id AND o.object_type_id = ot.object_type_id AND l.PREVIOUS_OBJECT_ID = m.MODULE_ID AND l.LINK_TYPE_ID = 3 
AND cs.OBJECT_ID = l.NEXT_OBJECT_ID AND cs.SNAPSHOT_ID in (select max(snapshot_id) from dss_snapshots) 
AND cs.CHANGE_TYPE = 2);"""):
           
            if line2 is not None:                
                comn = line2[0]
                if comn >= 100:
                    comn = round(comn, 2) 
                    # comn = line2[0]
                    worksheet.write(4, 1, comn, adtval_algn)
                    worksheet.write(4, 2, "FAILED", adtval_algn1)
                else:
                    # comn = line2[0]  
                    comn = round(comn, 2)    
                    worksheet.write(4, 1, comn, adtval_algn) 
                    worksheet.write(4, 2, "PASSED", adtval_algn) 
            
#-------------------TQI----------------------------------------------

        indx = 5
        worksheet.write(indx, 0, "TQI Variation (in %)", bold)
        worksheet.write(indx, 1, 'NA', adtval_algn) 
        worksheet.write(indx, 2, 'NA', adtval_algn)
        for line3 in central.execute_query("""select prev_snap.metric_name,
case when prev_snap.metric_num_value <> 0 then round(((curr_snap.metric_num_value - prev_snap.metric_num_value) * 100 / prev_snap.metric_num_value),2) else 0 end varPercent
from
(
select dmt.metric_id, dmt.metric_name, dmr.metric_num_value, metric_value_index
from dss_metric_results dmr, dss_metric_types dmt, dss_objects o
where
dmr.metric_id = dmt.metric_id
and dmr.object_id = o.object_id
and o.object_type_id = -102
and snapshot_id = (select max(snapshot_id) from dss_snapshots)
and
(dmr.metric_value_index = 0 and dmr.metric_id in (60017)))
curr_snap,
(
select dmt.metric_id, dmt.metric_name, dmr.metric_num_value, metric_value_index
from dss_metric_results dmr, dss_metric_types dmt, dss_objects o
where
dmr.metric_id = dmt.metric_id
and dmr.object_id = o.object_id
and o.object_type_id = -102
and snapshot_id = (select max(snapshot_id) from dss_snapshots where snapshot_id < (select max(snapshot_id) from dss_snapshots))
and (dmr.metric_value_index = 0 and dmr.metric_id in (60017)))prev_snap
where curr_snap.metric_id=prev_snap.metric_id
and curr_snap.metric_value_index=prev_snap.metric_value_index;"""): 
#             if line3[0] is not None:
#                 metrcname = line3[0]
            if line3[1] is not None:    
                varPercent = line3[1]
                if varPercent >= 5 or varPercent <= -5:
                    varPercent = round(varPercent, 2)
                    worksheet.write(indx, 1, str(varPercent) + '%', adtval_algn) 
                    worksheet.write(indx, 2, "FAILED", adtval_algn1)
                else:
                    varPercent = round(varPercent, 2)
                    worksheet.write(indx, 1, str(varPercent) + '%', adtval_algn) 
                    worksheet.write(indx, 2, "PASSED", adtval_algn)              
                       
            indx = indx + 1        
            
#-------------Code Lines Variation-------------------------------------
        
        indx = 6
        worksheet.write(indx, 0, 'Code Lines Variation (in %)', bold)
        worksheet.write(indx, 1, 'NA', adtval_algn) 
        worksheet.write(indx, 2, 'NA', adtval_algn)
        for line3 in central.execute_query("""select prev_snap.metric_name,
case when prev_snap.metric_num_value <> 0 then round(((curr_snap.metric_num_value - prev_snap.metric_num_value) * 100 / prev_snap.metric_num_value),2) else 0 end varPercent
 from
(
select dmt.metric_id, dmt.metric_name, dmr.metric_num_value, metric_value_index
from dss_metric_results dmr, dss_metric_types dmt, dss_objects o
where
dmr.metric_id = dmt.metric_id
and dmr.object_id = o.object_id
and o.object_type_id = -102
and snapshot_id = (select max(snapshot_id) from dss_snapshots)
and
(dmr.metric_value_index = 1 and dmr.metric_id in (10151)))
curr_snap,
(
select dmt.metric_id, dmt.metric_name, dmr.metric_num_value, metric_value_index
from dss_metric_results dmr, dss_metric_types dmt, dss_objects o
where
dmr.metric_id = dmt.metric_id
and dmr.object_id = o.object_id
and o.object_type_id = -102
and snapshot_id = (select max(snapshot_id) from dss_snapshots where snapshot_id < (select max(snapshot_id) from dss_snapshots))
and (dmr.metric_value_index = 1 and dmr.metric_id in (10151)))prev_snap
where curr_snap.metric_id=prev_snap.metric_id
and curr_snap.metric_value_index=prev_snap.metric_value_index;"""): 
#             if line3[0] is not None:
#                 metrcname = line3[0]                    
            if line3[1] is not None:    
                varPercent = line3[1]
                if varPercent >= 5 or varPercent <= -5:
                    varPercent = round(varPercent, 2)
                    worksheet.write(indx, 1, str(varPercent) + '%', adtval_algn) 
                    worksheet.write(indx, 2, "FAILED", adtval_algn1)
                else:
                    varPercent = round(varPercent, 2)
                    worksheet.write(indx, 1, str(varPercent) + '%', adtval_algn) 
                    worksheet.write(indx, 2, "PASSED", adtval_algn) 
                    
            indx = indx + 1 
            
#-------------Transferabilty (in %)-------------------------------------
        indx = 7
        worksheet.write(indx, 0, 'Transferabilty (in %)', bold)
        worksheet.write(indx, 1, 'NA', adtval_algn) 
        worksheet.write(indx, 2, 'NA', adtval_algn)
        for line3 in central.execute_query("""select prev_snap.metric_name,
case when prev_snap.metric_num_value <> 0 then round(((curr_snap.metric_num_value - prev_snap.metric_num_value) * 100 / prev_snap.metric_num_value),2) else 0 end varPercent
from
(
select dmt.metric_id, dmt.metric_name, dmr.metric_num_value, metric_value_index
from dss_metric_results dmr, dss_metric_types dmt, dss_objects o
where
dmr.metric_id = dmt.metric_id
and dmr.object_id = o.object_id
and o.object_type_id = -102
and snapshot_id = (select max(snapshot_id) from dss_snapshots)
and
(dmr.metric_value_index = 0 and dmr.metric_id in (60011)))
curr_snap,
(
select dmt.metric_id, dmt.metric_name, dmr.metric_num_value, metric_value_index
from dss_metric_results dmr, dss_metric_types dmt, dss_objects o
where dmr.metric_id = dmt.metric_id
and dmr.object_id = o.object_id
and o.object_type_id = -102
and snapshot_id = (select max(snapshot_id) from dss_snapshots where snapshot_id < (select max(snapshot_id) from dss_snapshots))
and (dmr.metric_value_index = 0 and dmr.metric_id in (60011)))prev_snap
where curr_snap.metric_id=prev_snap.metric_id
and curr_snap.metric_value_index=prev_snap.metric_value_index;"""): 
#             if line3[0] is not None:
#                 metrcname = line3[0]                    
            if line3[1] is not None:    
                varPercent = line3[1]
                if varPercent >= 5 or varPercent <= -5:
                    varPercent = round(varPercent, 2)
                    worksheet.write(indx, 1, str(varPercent) + '%', adtval_algn)  
                    worksheet.write(indx, 2, "FAILED", adtval_algn1)
                else:
                    varPercent = round(varPercent, 2)
                    worksheet.write(indx, 1, str(varPercent) + '%', adtval_algn)  
                    worksheet.write(indx, 2, "PASSED", adtval_algn) 
           
            else:
                worksheet.write(indx, 1, 'NA', adtval_algn) 
                worksheet.write(indx, 2, 'NA', adtval_algn)  
                        
            indx = indx + 1 

#-------------Efficiency (in %)-------------------------------------
        
        indx = 8
        worksheet.write(indx, 0, 'Efficiency (in %)', bold)
        worksheet.write(indx, 1, 'NA', adtval_algn) 
        worksheet.write(indx, 2, 'NA', adtval_algn)
        for line3 in central.execute_query("""select prev_snap.metric_name,
case when prev_snap.metric_num_value <> 0 then round(((curr_snap.metric_num_value - prev_snap.metric_num_value) * 100 / prev_snap.metric_num_value),2) else 0 end varPercent
from
(
select dmt.metric_id, dmt.metric_name, dmr.metric_num_value, metric_value_index
from dss_metric_results dmr, dss_metric_types dmt, dss_objects o
where
dmr.metric_id = dmt.metric_id
and dmr.object_id = o.object_id
and o.object_type_id = -102
and snapshot_id = (select max(snapshot_id) from dss_snapshots)
and
(dmr.metric_value_index = 0 and dmr.metric_id in (60014)))
curr_snap,
(
select dmt.metric_id, dmt.metric_name, dmr.metric_num_value, metric_value_index
from dss_metric_results dmr, dss_metric_types dmt, dss_objects o
where
dmr.metric_id = dmt.metric_id
and dmr.object_id = o.object_id
and o.object_type_id = -102
and snapshot_id = (select max(snapshot_id) from dss_snapshots where snapshot_id < (select max(snapshot_id) from dss_snapshots))
and (dmr.metric_value_index = 0 and dmr.metric_id in (60014)))prev_snap
where curr_snap.metric_id=prev_snap.metric_id
and curr_snap.metric_value_index=prev_snap.metric_value_index;"""): 
#             if line3[0] is not None:
#                 metrcname = line3[0]                    
            if line3[1] is not None:    
                varPercent = line3[1]
                if varPercent >= 5 or varPercent <= -5:
                    varPercent = round(varPercent, 2) 
                    worksheet.write(indx, 1, str(varPercent) + '%', adtval_algn)
                    worksheet.write(indx, 2, "FAILED", adtval_algn1)
                else:
                    varPercent = round(varPercent, 2) 
                    worksheet.write(indx, 1, str(varPercent) + '%', adtval_algn)
                    worksheet.write(indx, 2, "PASSED", adtval_algn) 
            else:
                worksheet.write(indx, 1, 'NA', adtval_algn)  
                worksheet.write(indx, 2, 'NA', adtval_algn)       
                      
            indx = indx + 1
         
#-------------Security (in %)-------------------------------------
        
        indx = 9
        worksheet.write(indx, 0, 'Security (in %)', bold)
        worksheet.write(indx, 1, 'NA', adtval_algn) 
        worksheet.write(indx, 2, 'NA', adtval_algn)
        for line3 in central.execute_query("""select prev_snap.metric_name,
case when prev_snap.metric_num_value <> 0 then round(((curr_snap.metric_num_value - prev_snap.metric_num_value) * 100 / prev_snap.metric_num_value),2) else 0 end varPercent
from
(
select dmt.metric_id, dmt.metric_name, dmr.metric_num_value, metric_value_index
from dss_metric_results dmr, dss_metric_types dmt, dss_objects o
where
dmr.metric_id = dmt.metric_id
and dmr.object_id = o.object_id
and o.object_type_id = -102
and snapshot_id = (select max(snapshot_id) from dss_snapshots)
and
(dmr.metric_value_index = 0 and dmr.metric_id in (60016)))
curr_snap,
(
select dmt.metric_id, dmt.metric_name, dmr.metric_num_value, metric_value_index
from dss_metric_results dmr, dss_metric_types dmt, dss_objects o
where
dmr.metric_id = dmt.metric_id
and dmr.object_id = o.object_id
and o.object_type_id = -102
and snapshot_id = (select max(snapshot_id) from dss_snapshots where snapshot_id < (select max(snapshot_id) from dss_snapshots))
and (dmr.metric_value_index = 0 and dmr.metric_id in (60016)))prev_snap
where curr_snap.metric_id=prev_snap.metric_id
and curr_snap.metric_value_index=prev_snap.metric_value_index;"""): 
#             if line3[0] is not None:
#                 metrcname = line3[0]                    
            if line3[1] is not None:    
                varPercent = line3[1]
                if varPercent >= 5 or varPercent <= -5:
                    varPercent = round(varPercent, 2)
                    worksheet.write(indx, 1, str(varPercent) + '%', adtval_algn)
                    worksheet.write(indx, 2, "FAILED", adtval_algn1)
                else:
                    varPercent = round(varPercent, 2)
                    worksheet.write(indx, 1, str(varPercent) + '%', adtval_algn)
                    worksheet.write(indx, 2, "PASSED", adtval_algn) 
                                
            else:
                worksheet.write(indx, 1, 'NA', adtval_algn)  
                worksheet.write(indx, 2, 'NA', adtval_algn)
                      
            indx = indx + 1
         
#-------------Changeability(in %)-------------------------------------
        
        indx = 10
        worksheet.write(indx, 0, 'Changeability (in %)', bold)
        worksheet.write(indx, 1, 'NA', adtval_algn) 
        worksheet.write(indx, 2, 'NA', adtval_algn)
        for line3 in central.execute_query("""select prev_snap.metric_name,
case when prev_snap.metric_num_value <> 0 then round(((curr_snap.metric_num_value - prev_snap.metric_num_value) * 100 / prev_snap.metric_num_value),2) else 0 end varPercent
from
(
select dmt.metric_id, dmt.metric_name, dmr.metric_num_value, metric_value_index
from dss_metric_results dmr, dss_metric_types dmt, dss_objects o
where
dmr.metric_id = dmt.metric_id
and dmr.object_id = o.object_id
and o.object_type_id = -102
and snapshot_id = (select max(snapshot_id) from dss_snapshots)
and
(dmr.metric_value_index = 0 and dmr.metric_id in (60012)))
curr_snap,
(
select dmt.metric_id, dmt.metric_name, dmr.metric_num_value, metric_value_index
from dss_metric_results dmr, dss_metric_types dmt, dss_objects o
where
dmr.metric_id = dmt.metric_id
and dmr.object_id = o.object_id
and o.object_type_id = -102
and snapshot_id = (select max(snapshot_id) from dss_snapshots where snapshot_id < (select max(snapshot_id) from dss_snapshots))
and (dmr.metric_value_index = 0 and dmr.metric_id in (60012)))prev_snap
where curr_snap.metric_id=prev_snap.metric_id
and curr_snap.metric_value_index=prev_snap.metric_value_index;"""): 
#             if line3[0] is not None:
#                 metrcname = line3[0]                    
            if line3[1] is not None:    
                varPercent = line3[1]
                if varPercent >= 5 or varPercent <= -5:
                    varPercent = round(varPercent, 2) 
                    worksheet.write(indx, 1, str(varPercent) + '%', adtval_algn) 
                    worksheet.write(indx, 2, "FAILED", adtval_algn1)
                else:
                    varPercent = round(varPercent, 2) 
                    worksheet.write(indx, 1, str(varPercent) + '%', adtval_algn) 
                    worksheet.write(indx, 2, "PASSED", adtval_algn) 
                             
            else:
                worksheet.write(indx, 1, 'NA', adtval_algn)  
                worksheet.write(indx, 2, 'NA', adtval_algn)
                     
            indx = indx + 1
#-------------Robustness(in %)-------------------------------------
        indx = 11
        worksheet.write(indx, 0, 'Robustness (in %)', bold)
        worksheet.write(indx, 1, 'NA', adtval_algn) 
        worksheet.write(indx, 2, 'NA', adtval_algn)
        for line3 in central.execute_query("""select prev_snap.metric_name,
case when prev_snap.metric_num_value <> 0 then round(((curr_snap.metric_num_value - prev_snap.metric_num_value) * 100 / prev_snap.metric_num_value),2) else 0 end varPercent
from
(
select dmt.metric_id, dmt.metric_name, dmr.metric_num_value, metric_value_index
from dss_metric_results dmr, dss_metric_types dmt, dss_objects o
where
dmr.metric_id = dmt.metric_id
and dmr.object_id = o.object_id
and o.object_type_id = -102
and snapshot_id = (select max(snapshot_id) from dss_snapshots)
and
(dmr.metric_value_index = 0 and dmr.metric_id in (60013)))
curr_snap,
(
select dmt.metric_id, dmt.metric_name, dmr.metric_num_value, metric_value_index
from dss_metric_results dmr, dss_metric_types dmt, dss_objects o
where
dmr.metric_id = dmt.metric_id
and dmr.object_id = o.object_id
and o.object_type_id = -102
and snapshot_id = (select max(snapshot_id) from dss_snapshots where snapshot_id < (select max(snapshot_id) from dss_snapshots))
and (dmr.metric_value_index = 0 and dmr.metric_id in (60013)))prev_snap
where curr_snap.metric_id=prev_snap.metric_id
and curr_snap.metric_value_index=prev_snap.metric_value_index;"""): 
#             if line3[0] is not None:
#                 metrcname = line3[0]                    
            if line3[1] is not None:    
                varPercent = line3[1]
                if varPercent >= 5 or varPercent <= -5:
                    varPercent = round(varPercent, 2)   
                    worksheet.write(indx, 1, str(varPercent) + '%', adtval_algn) 
                    worksheet.write(indx, 2, "FAILED", adtval_algn1)
                else:
                    varPercent = round(varPercent, 2) 
                    worksheet.write(indx, 1, str(varPercent) + '%', adtval_algn)   
                    worksheet.write(indx, 2, "PASSED", adtval_algn) 
                    
            indx = indx + 1
#-------------Critical Violation Variation (in %)-------------------------------------
        
        indx = 12
        worksheet.write(indx, 0, 'Critical Violation Variation (in %)', bold)
        worksheet.write(indx, 1, 'NA', adtval_algn) 
        worksheet.write(indx, 2, 'NA', adtval_algn)
        for line3 in central.execute_query("""select prev_snap.metric_name,
case when prev_snap.metric_num_value <> 0 then round(((curr_snap.metric_num_value - prev_snap.metric_num_value) * 100 / prev_snap.metric_num_value),2) else 0 end varPercent
 from
(
select dmt.metric_id, dmt.metric_name, dmr.metric_num_value, metric_value_index
from dss_metric_results dmr, dss_metric_types dmt, dss_objects o
where 
dmr.metric_id = dmt.metric_id 
and dmr.object_id = o.object_id
and o.object_type_id = -102
and snapshot_id = (select max(snapshot_id) from dss_snapshots)
and 
(
dmr.metric_value_index = 1 and dmr.metric_id in (67011)))
curr_snap,
(
select dmt.metric_id, dmt.metric_name, dmr.metric_num_value, metric_value_index
from dss_metric_results dmr, dss_metric_types dmt, dss_objects o
where 
dmr.metric_id = dmt.metric_id 
and dmr.object_id = o.object_id
and o.object_type_id = -102
and snapshot_id = (select max(snapshot_id) from dss_snapshots where snapshot_id < (select max(snapshot_id) from dss_snapshots))
and (
dmr.metric_value_index = 1 and dmr.metric_id in (67011)))prev_snap
where curr_snap.metric_id=prev_snap.metric_id
and curr_snap.metric_value_index=prev_snap.metric_value_index;"""):
#             if line3[0] is not None:
#                 metrcname = line3[0]                    
            if line3[1] is not None:    
                varPercent = line3[1]
                if varPercent >= 5 or varPercent <= -5:
                    varPercent = round(varPercent, 2)   
                    worksheet.write(indx, 1, str(varPercent) + '%', adtval_algn) 
                    worksheet.write(indx, 2, "FAILED", adtval_algn1)
                else:
                    varPercent = round(varPercent, 2) 
                    worksheet.write(indx, 1, str(varPercent) + '%', adtval_algn)   
                    worksheet.write(indx, 2, "PASSED", adtval_algn) 
                          
            else:
                worksheet.write(indx, 1, 'NA', adtval_algn)  
                worksheet.write(indx, 2, 'NA', adtval_algn)
                    
            indx = indx + 1
            
#-------------Added Quality Rule-------------------------------------

        indx = 13
        worksheet.write(indx, 0, 'Added Quality Rule', bold)
        worksheet.write(indx, 1, 'NA', adtval_algn) 
        worksheet.write(indx, 2, 'NA', adtval_algn)
        for line3 in central.execute_query("""select count(SNAP3.METRIC_ID)
from (select distinct(metric_id) from dss_metric_results where snapshot_id=(select max(snapshot_id) from dss_snapshots)) SNAP3
where SNAP3.METRIC_ID not in (select distinct(metric_id) from dss_metric_results where snapshot_id=(select snapshot_id from (select snapshot_id from dss_snapshots order by snapshot_id desc limit 2) as snapid order by 
snapshot_id limit 1)) and SNAP3.METRIC_ID%2 = 0 and SNAP3.metric_id not in (10356,10358);"""):  
            if line3[0] is not None:
                cnt = line3[0]           
                if cnt > 0:
                    worksheet.write(indx, 1, cnt, adtval_algn)   
                    worksheet.write(indx, 2, "FAILED", adtval_algn1)
                else:
                    worksheet.write(indx, 1, cnt, adtval_algn)   
                    worksheet.write(indx, 2, "PASSED", adtval_algn) 
            else:
                worksheet.write(indx, 1, 'NA', adtval_algn)  
                worksheet.write(indx, 2, 'NA', adtval_algn)
                     
            indx = indx + 1
            
#-------------Deleted Quality Rule-------------------------------------  
        indx = 14
        worksheet.write(indx, 0, 'Deleted Quality Rule', bold)
        worksheet.write(indx, 1, 'NA', adtval_algn) 
        worksheet.write(indx, 2, 'NA', adtval_algn)
        for line3 in central.execute_query("""select count(SNAP2.METRIC_ID)
from (select distinct(metric_id) from dss_metric_results where snapshot_id=(select snapshot_id from (select snapshot_id from dss_snapshots order by snapshot_id desc limit 2) as snapid order by snapshot_id limit 1)) SNAP2
where SNAP2.METRIC_ID not in (select distinct(metric_id) from dss_metric_results where snapshot_id=(select max(snapshot_id) from dss_snapshots))
and SNAP2.METRIC_ID%2 = 0 and SNAP2.METRIC_ID not in (10356,10358);"""):       
            if line3[0] is not None:
                lst_cnt = line3[0]             
                if lst_cnt > 0:
                    worksheet.write(indx, 1, lst_cnt, adtval_algn) 
                    worksheet.write(indx, 2, "FAILED", adtval_algn1)
                else:
                    worksheet.write(indx, 1, lst_cnt, adtval_algn) 
                    worksheet.write(indx, 2, "PASSED", adtval_algn) 
            else:
                worksheet.write(indx, 1, 'NA', adtval_algn)  
                worksheet.write(indx, 2, 'NA', adtval_algn)
                      
            indx = indx + 1

#-------------AFP variation in %-------------------------------------  
        indx = 15
        worksheet.write(indx, 0, 'AFP variation in %', bold)
        worksheet.write(indx, 1, 'NA', adtval_algn) 
        worksheet.write(indx, 2, 'NA', adtval_algn)
        for line3 in central.execute_query("""select prev_snap.metric_name,
case when prev_snap.metric_num_value <> 0 then round(((curr_snap.metric_num_value - prev_snap.metric_num_value) * 100 / prev_snap.metric_num_value),2) else 0 end varPercent
 from
(
select dmt.metric_id, dmt.metric_name, dmr.metric_num_value, metric_value_index
from dss_metric_results dmr, dss_metric_types dmt, dss_objects o
where 
dmr.metric_id = dmt.metric_id 
and dmr.object_id = o.object_id
and o.object_type_id = -102
and snapshot_id = (select max(snapshot_id) from dss_snapshots)
and 
(
dmr.metric_value_index = 1 and dmr.metric_id in (10202)))
curr_snap,
(
select dmt.metric_id, dmt.metric_name, dmr.metric_num_value, metric_value_index
from dss_metric_results dmr, dss_metric_types dmt, dss_objects o
where 
dmr.metric_id = dmt.metric_id 
and dmr.object_id = o.object_id
and o.object_type_id = -102
and snapshot_id = (select max(snapshot_id) from dss_snapshots where snapshot_id < (select max(snapshot_id) from dss_snapshots))
and ( 
dmr.metric_value_index = 1 and dmr.metric_id in (10202)))prev_snap
where curr_snap.metric_id=prev_snap.metric_id
and curr_snap.metric_value_index=prev_snap.metric_value_index;"""):
#             if line3[0] is not None:
#                 metrcname = line3[0]                    
            if line3[1] is not None:    
                varPercent = line3[1]
                if varPercent >= 5 or varPercent <= -5:
                    varPercent = round(varPercent, 2) 
                    worksheet.write(indx, 1, str(varPercent) + '%', adtval_algn)
                    worksheet.write(indx, 2, "FAILED", adtval_algn1)
                else:
                    varPercent = round(varPercent, 2) 
                    worksheet.write(indx, 1, str(varPercent) + '%', adtval_algn)
                    worksheet.write(indx, 2, "PASSED", adtval_algn)        
                                            
            else:
                worksheet.write(indx, 1, 'NA', adtval_algn)  
                worksheet.write(indx, 2, 'NA', adtval_algn)
                      
            indx = indx + 1
            
#-------------AEFP/AETP Ratio for latest 2 snapshots-------------------------------------
        indx = 16
        worksheet.write(indx, 0, 'AEFP/AETP ratio for latest 2 snapshot', bold)
        worksheet.write(indx, 1, 'NA', adtval_algn) 
        worksheet.write(indx, 2, 'NA', adtval_algn)
        for line3 in central.execute_query("""select(
(select
case when c2.metric_num_value <> 0 then round((c1.metric_num_value/ c2.metric_num_value)*100,2)   else 0 end "Current Ratio" from 
dss_metric_results c1, dss_metric_results c2
where c1.metric_id  = 10430 and c2.metric_id = 10440 
and c1.object_id in (select object_id from dss_objects where object_type_id =-102) 
and c2.object_id in (select object_id from dss_objects where object_type_id =-102) 
and c1.snapshot_id = (select max(snapshot_id) from dss_snapshots)
and c2.snapshot_id = (select max(snapshot_id) from dss_snapshots) )
-
(select
case when c4.metric_num_value <> 0 then round((c3.metric_num_value/ nullif(c4.metric_num_value,0))*100,2)  else 0 end "Previous Ratio" from 
dss_metric_results c3, dss_metric_results c4
where c3.metric_id  = 10430 and c4.metric_id = 10440 
and c3.snapshot_id = (select snapshot_id from (select snapshot_id from dss_snapshots order by snapshot_id desc limit 2) as snapid order by snapshot_id limit 1)
and c3.object_id in (select object_id from dss_objects where object_type_id =-102) 
and c4.object_id in (select object_id from dss_objects where object_type_id =-102) 
and c4.snapshot_id = (select snapshot_id from (select snapshot_id from dss_snapshots order by snapshot_id desc limit 2) as snapid order by snapshot_id limit 1)))
as "AEFP/AETP Ratio Difference Btw last 2 snapshots";"""):
            if line3[0] is not None:
                varPercent = line3[0]           
                if varPercent <= -5:
                    varPercent = round(varPercent, 2) 
                    worksheet.write(indx, 1, varPercent, adtval_algn)
                    worksheet.write(indx, 2, "FAILED", adtval_algn1)
                else:
                    varPercent = round(varPercent, 2) 
                    worksheet.write(indx, 1, varPercent, adtval_algn)
                    worksheet.write(indx, 2, "PASSED", adtval_algn) 
                                                    
            else:
                worksheet.write(indx, 1, 'NA', adtval_algn)  
                worksheet.write(indx, 2, 'NA', adtval_algn)
                   
            indx = indx + 1
 
#-------------LOC/FP Ratio-------------------------------------    
        indx = 17
        worksheet.write(indx, 0, 'LOC/FP Ratio', bold)
        worksheet.write(indx, 1, 'NA', adtval_algn) 
        worksheet.write(indx, 2, 'NA', adtval_algn)
        for line3 in central.execute_query("""select(
(select
case when c1.metric_num_value <> 0 then round((c2.metric_num_value / c1.metric_num_value),2) else 0 end "CurrentRatio"
from dss_metric_results c1, dss_metric_results c2
where c1.metric_id  = 10202 and c2.metric_id = 10151 
and c1.object_id in (select object_id from dss_objects where object_type_id =-102) 
and c2.object_id in (select object_id from dss_objects where object_type_id =-102) 
and c1.snapshot_id = (select max(snapshot_id) from dss_snapshots)
and c2.snapshot_id = (select max(snapshot_id) from dss_snapshots) )

-
(select case when c3.metric_num_value <> 0 then round((c4.metric_num_value / c3.metric_num_value),2) else 0 end "PreviousRatio"
from dss_metric_results c3, dss_metric_results c4
where  c3.metric_id  = 10202 and c4.metric_id = 10151 
and c3.snapshot_id = (select snapshot_id from (select snapshot_id from dss_snapshots order by snapshot_id desc limit 2) as snapid order by 
                                                                          snapshot_id limit 1)
                                                                          and c3.object_id in (select object_id from dss_objects where object_type_id =-102) 
and c4.object_id in (select object_id from dss_objects where object_type_id =-102) 
and c4.snapshot_id = (select snapshot_id from (select snapshot_id from dss_snapshots order by snapshot_id desc limit 2) as snapid order by 
                                                                          snapshot_id limit 1)));"""):
            if line3[0] is not None:
                varPercent = line3[0]           
                if varPercent <= -5:
                    varPercent = round(varPercent, 2)
                    worksheet.write(indx, 1, varPercent, adtval_algn)
                    worksheet.write(indx, 2, "FAILED", adtval_algn1)
                else:
                    varPercent = round(varPercent, 2)
                    worksheet.write(indx, 1, varPercent, adtval_algn)
                    worksheet.write(indx, 2, "PASSED", adtval_algn)    
                                                 
            else:
                worksheet.write(indx, 1, 'NA', adtval_algn)  
                worksheet.write(indx, 2, 'NA', adtval_algn)
                       
            indx = indx + 1
#-------------Artifacts Coverage Variation %------------------------------------- 
        indx = 18
        worksheet.write(indx, 0, 'Artifacts Coverage Variation %', bold)
        worksheet.write(indx, 1, 'NA', adtval_algn)  
        worksheet.write(indx, 2, 'NA', adtval_algn)
        for line3 in central.execute_query("""select case when c2.percentage::integer <> 0 then round(((c1.percentage::integer - c2.percentage::integer) / c2.percentage::integer))*100 else 0 end "Artefact Coverage Variation %"
from aia_console.smelltest_history c1,aia_console.smelltest_history c2    where c1.currentversion = (select snapshot_name from dss_snapshots where snapshot_id in (select max(snapshot_id) from dss_snapshots))    
and c2.currentversion = (select snapshot_name from dss_snapshots where snapshot_id in ((select snapshot_id from (select snapshot_id from dss_snapshots order by snapshot_id desc limit 2) as snapid order by 
snapshot_id limit 1))) and c1.checknumber = '1smell' and c2.checknumber = '1smell' and c1.appname = c2.appname;"""):
              
            if line3[0] is not None:
                varPercent = line3[0]           
                if varPercent <= -3:
                    varPercent = round(varPercent, 2) 
                    worksheet.write(indx, 1, str(varPercent) + '%', adtval_algn)   
                    worksheet.write(indx, 2, "FAILED", adtval_algn1)
                else:
                    varPercent = round(varPercent, 2) 
                    worksheet.write(indx, 1, str(varPercent) + '%', adtval_algn)   
                    worksheet.write(indx, 2, "PASSED", adtval_algn)    
            
#             elif line3[0] == '':  
#                 worksheet.write(indx, 1, 'NA', adtval_algn)  
#                 worksheet.write(indx, 2, 'NA', adtval_algn)
                        
            indx = indx + 1
#-------------Empty Transaction Variation %-------------------------------------           
        indx = 19
        worksheet.write(indx, 0, 'Empty Transaction Variation %', bold)
        worksheet.write(indx, 1, 'NA', adtval_algn) 
        worksheet.write(indx, 2, 'NA', adtval_algn)
        for line3 in central.execute_query("""select distinct case when c2.percentage::integer <> 0 then round(((c1.percentage::integer - c2.percentage::integer) / c2.percentage::integer))*100 else 0 end "Empty Transaction Variation %"
from aia_console.smelltest_history c1,aia_console.smelltest_history c2    where c1.currentversion = (select snapshot_name from dss_snapshots where snapshot_id in (select max(snapshot_id) from dss_snapshots))    
and c2.currentversion = (select snapshot_name from dss_snapshots where snapshot_id in ((select snapshot_id from (select snapshot_id from dss_snapshots order by snapshot_id desc limit 2) as snapid order by 
snapshot_id limit 1))) and c1.checknumber = '2smell' and c2.checknumber = '2smell' and c1.appname = c2.appname;"""):
            if line3[0] is not None:
                varPercent = line3[0]           
                if varPercent >= 3:
                    varPercent = round(varPercent, 2)
                    worksheet.write(indx, 1, str(varPercent) + '%', adtval_algn) 
                    worksheet.write(indx, 2, "FAILED", adtval_algn1)
                else:
                    varPercent = round(varPercent, 2)
                    worksheet.write(indx, 1, str(varPercent) + '%', adtval_algn) 
                    worksheet.write(indx, 2, "PASSED", adtval_algn)                                  
            
            else:
                worksheet.write(indx, 1, 'NA', adtval_algn)  
                worksheet.write(indx, 2, 'NA', adtval_algn)
                       
            indx = indx + 1
#-------------%added code not part of AFP------------------------------------- 
        
        indx = 20
        worksheet.write(indx, 0, '%added code not part of AFP', bold)
        worksheet.write(indx, 1, 'NA', adtval_algn) 
        worksheet.write(indx, 2, 'NA', adtval_algn)
        for line3 in central.execute_query("""SELECT CASE WHEN TA."TOTAL ADDED ARTIFACTS" <> 0 THEN round(("ARTIFACTS_NOT_CONTRIBUTING_TO_FP"*100)/TA."TOTAL ADDED ARTIFACTS") ELSE 0 END AS "%_ARTIFACTS_NOT_CONTRIBUTING_TO_FP"
FROM
(SELECT Count(1) as "TOTAL ADDED ARTIFACTS"
FROM CSV_OBJECTS_STATUSES COS
WHERE  COS.SNAPHOT_ID = (SELECT MAX(SNAPSHOT_ID) FROM DSS_SNAPSHOTS)
AND COS.OBJECT_STATUS = 'Added'
AND COS.OBJECT_IS_ARTIFACT = 'Artifact') TA,
(SELECT COUNT(1) AS "ARTIFACTS_NOT_CONTRIBUTING_TO_FP"
FROM AEP_TECHNICAL_ARTIFACTS_VW A JOIN DSS_OBJECTS O 
ON A.OBJECT_ID = O.OBJECT_ID                                                                                               
JOIN DSS_METRIC_RESULTS DC
ON  A.SNAPSHOT_ID = DC.SNAPSHOT_ID                                                                             
AND DC.METRIC_ID = 10359                                                                                                       
JOIN DSS_SNAPSHOTS S 
ON A.SNAPSHOT_ID = (select max(snapshot_id) from dss_snapshots)
AND S.SNAPSHOT_ID = A.SNAPSHOT_ID 
AND  S.SNAPSHOT_STATUS = 2
AND S.ENHANCEMENT_MEASURE::TEXT = 'AEP'::TEXT                    
JOIN DSS_SNAPSHOT_INFO SI 
ON S.SNAPSHOT_ID = SI.SNAPSHOT_ID 
AND SI.object_id = S.application_id                  
WHERE A.STATUS = 'ADDED'                                                                                                      
AND DC.METRIC_VALUE_INDEX = 1) ANCTF;"""):
            if line3[0] is not None:
                varPercent = line3[0]           
                if varPercent >= 5:
                    varPercent = round(varPercent, 2)
                    worksheet.write(indx, 1, str(varPercent) + '%', adtval_algn)
                    worksheet.write(indx, 2, "FAILED", adtval_algn1)
                else:
                    varPercent = round(varPercent, 2)
                    worksheet.write(indx, 1, str(varPercent) + '%', adtval_algn)
                    worksheet.write(indx, 2, "PASSED", adtval_algn)                                  
            else:
                worksheet.write(indx, 1, 'NA', adtval_algn)  
                worksheet.write(indx, 2, 'NA', adtval_algn)
                   
            indx = indx + 1
 
#-------------Technologies not interacting with other technologies-------------------------------------             
        
        indx = 23
        worksheet.write(indx, 0, 'Technologies not interacting with other technologies', ylclr)       
        indx = indx + 1
        for line3 in kb.execute_query("""select distinct object_language_name from cdt_objects where 
object_language_name not like '%N/A%' 
and lower(object_type_str) not like lower('% subset')
and object_type_str not like '% Directory'
and object_type_str not like '% Project%'
and object_language_name not in 
(select object_language_name from cdt_objects where object_id in 
(select c1.idclr from 
acc c1 
join cdt_objects c2 
on c1.idclr = c2.object_id 
join cdt_objects c3
on c3.object_id = c1.idcle 
where c3.object_language_name != c2.object_language_name
union
select c1.idcle from acc c1
join cdt_objects c2 
on c1.idclr = c2.object_id 
join cdt_objects c3
on c3.object_id = c1.idcle 
where c3.object_language_name != c2.object_language_name));"""):
            
            if line3[0] is not None:
                techName = line3[0]               
                worksheet.write(indx, 0, techName, adtval_algn)            
                indx = indx + 1
#-------------Transaction with Large SCC Group------------------------------------- 
        indx = 25
        worksheet.write(23, 5, 'Transaction with Large SCC Group', ylclr)
        worksheet.write(24, 4, 'Transaction name', ylclr)
        worksheet.write(24, 5, 'Object Type', ylclr)
        worksheet.write(24, 6, 'DET', ylclr)
        worksheet.write(24, 7, 'FTR', ylclr)
        indx = 23
        
        for line3 in kb.execute_query("""SELECT distinct trans_obj_id , trans_name , trans_type , dt.det, dt.ftr
from (
SELECT
t51.object_id       trans_obj_id,
t51.object_name     trans_name,
t51.object_type_str trans_type,
t31.object_id  art_id,
t31.object_name     art_name,
t31.object_type_Str art_type
FROM dss_objects t11,
dss_links t21,cdt_objects t31,dss_transaction t41,cdt_objects t51
WHERE
t11.object_type_id = 30002
AND t11.object_id = t21.previous_object_id
AND t21.next_object_id = t31.object_id
and t41.object_id = t11.object_id
and t41.form_id = t51.object_id
ORDER BY 1) A, dss_transaction dt
where dt.form_id = A.trans_obj_id
and A.art_id in (select object_id from fp_largescc)
and dt.cal_flags in (0,2);"""):
            if line3[1] is not None:
                scc_num = line3[1]
                worksheet.write(indx, 4, scc_num, adtval_algn)
            if line3[2] is not None:
                object_id = line3[2]    
                worksheet.write(indx, 5, object_id, adtval_algn)
            if line3[3] is not None:
                artfcts = line3[3]        
                worksheet.write(indx, 6, artfcts, adtval_algn)
            if line3[4] is not None:
                aded = line3[4]     
                worksheet.write(indx, 7, aded, adtval_algn)
            indx = indx + 1
                    
#-------------------Technology Report------------------                
        worksheet = workbook.add_worksheet('Technology and Frameworks') 
        worksheet1 = workbook.add_worksheet('Validate Extension') 
          
        self.identifySupportedLanguages(mngt, 'postgres', 'localhost', 2280, worksheet, worksheet1)
        # worksheet = workbook.add_worksheet('Technology') 
        # self.compareVersionsOfSupExtensions(analyzerExtensionsMap, installedExtenMap, worksheet, indx)
        workbook.close()
        status = "OK"
        publish_report('RescanAuditReport', status, "Audit Rescan", '', detail_report_path=report_path)
        
    def getDeliveryReportPath(self, mngtName, srcDB, srcHost, srcPort):
        deliveryPath = ''
        applicationPath = ''
        versionPath = ''
        # cursor, conn = self.createConnection(srcDB, srcHost, srcPort)
        
        cursor = mngtName.execute_query(
            "select serverpath from cms_pref_sources;")
        for eachRow in cursor:
            deliveryPath = eachRow[0] + "\\data"
    
        cursor = mngtName.execute_query(
            "select '{' || replace(field_value,'uuid:','') || '}' from cms_dynamicfields where entity_guid ~* 'pmcportfolio.Application' and field_guid = 'entry'")
        for eachRow in cursor:
            applicationPath = "\\" + eachRow[0]
    
        cursor = mngtName.execute_query("select '{' || replace(field_value,'dmtid:','') || '}' from cms_dynamicfields where entity_guid ~* 'pmcdeliveryunit.ApplicationVersion' and object_id = (select max(object_id) from cms_dynamicfields where entity_guid ~* 'pmcdeliveryunit.ApplicationVersion')")
        for eachRow in cursor:
            versionPath = "\\" + eachRow[0] + "\\DmtDeliveryReport.xml"
    
        return deliveryPath + applicationPath + versionPath
       
    def identifySupportedLanguages(self, mngtName, srcDB, srcHost, srcPort, worksheet, worksheet1):
      
        identifiedLanguages = {}
        unsupportedExtensions = {}
        identifiedFrameworks = {}
        identifiedAnalyzers = {}
        frameWorkExtensionsMap = {}
        deliveryReportPath = self.getDeliveryReportPath(mngtName, srcDB, srcHost, srcPort)
        # deliveryReportPath = "C:\\Users\\KSH\\Documents\\My Received Files\\DMTDeliveryReport.xml"
        if os.path.isfile(deliveryReportPath):
            tree = ET.parse(deliveryReportPath)
            root = tree.getroot()
            for fileLang in root.iter('FileLanguage'):
                if int(fileLang.get('total')) > 0:
                    identifiedLanguages[fileLang.get('languageId')] = fileLang.get('total')
    
            for fileLang in root.iter('UnSupportedExtension'):
                if int(fileLang.get('total')) > 0:
                    unsupportedExtensions[fileLang.get('extensionId')] = fileLang.get('total')
    
            for fileLang in root.iter('Framework'):
                identifiedFrameworks[fileLang.get('frameworkId')] = str(fileLang.get('version'))
    
            for fileLang in root.iter('Analyzer'):
                identifiedAnalyzers[fileLang.get('analyzerName')] = str(fileLang.get('fileExtension'))
    
        else:
            worksheet.write(0, 1, "Delivery report Path doesn't exist")
    
        frameWorkExtensionsMap = self.extractLatestExtensions(identifiedFrameworks)
        analyzerExtensionsMap = self.extractLatestExtensions(identifiedAnalyzers)
        # print(frameWorkExtensionsMap)
        installedExtenMap = self.getInstalledExtnsInDB(mngtName, srcDB, srcHost, srcPort)
        # print(installedExtenMap)
        worksheet.set_column(0, 2, 40)
        indx = 0
        worksheet.write(indx, 1, "BELOW ARE THE UNSUPPORTED TECHNOLOGIES/FRAMEWORKS IDENTIFIED IN THIS APPLICATION")
        indx = indx + 1
        worksheet.write(indx, 0, "FRAMEWORKS")
        indx = indx + 1
        indx = self.printUnsupportedTechnologies(frameWorkExtensionsMap, worksheet, indx)
        indx = indx + 1
        worksheet.write(indx, 0, "TECHNOLOGIES")
        indx = indx + 1
        indx = self.printUnsupportedTechnologies(analyzerExtensionsMap, worksheet, indx)
        indx = indx + 1
        analyzerExtensionsMap.update(frameWorkExtensionsMap)
        indx = indx + 1        
        # worksheet1.write(indx, 0, " VALIDATING THE LATEST EXTENSION versus INSTALLED extension ")
        # indx = indx + 1
        self.compareVersionsOfSupExtensions(analyzerExtensionsMap, installedExtenMap, worksheet1, indx)

    # def compareVersionsOfSupExtensions(self, analyzerExtensionsMap, installedExtenMap, worksheet, indx):
    def compareVersionsOfSupExtensions(self, analyzerExtensionsMap, installedExtenMap, worksheet, indx):
        worksheet.set_column(0, 2, 40)
        worksheet.write(0, 0, " VALIDATING THE LATEST EXTENSION versus INSTALLED extension ")
        indx = 1
        for k in analyzerExtensionsMap:
            if analyzerExtensionsMap[k] != 'Unsupported Technology':
                extensionVersion = analyzerExtensionsMap[k]
                extVerTuple = str(extensionVersion).split("--")
                flagver = 0
                for key in installedExtenMap:
                    if str(key).lower() == str(extVerTuple[0]).lower():
                        if installedExtenMap[key] != extVerTuple[1]:
                            worksheet.write(indx, 1, " For " + k + " Framework/Technology latest available version is " + extVerTuple[1] + " but currently " + installedExtenMap[key] + " is installed")
                            flagver = 1
                            indx = indx + 1
                if flagver == 0:
                    worksheet.write(indx, 1, " For " + k + " Framework/Technology latest available version is " + extVerTuple[1] + " but no extension is installed")
                    indx = indx + 1

    def printUnsupportedTechnologies(self, artifactMap, worksheet, idx):
        # idx = 2
        for k in artifactMap:
            if artifactMap[k] == 'Unsupported Technology':
                worksheet.write(idx, 1, k)
                idx = idx + 1
        return idx

    def extractLatestExtensions(self, identifiedartifacts):
        # print(type(identifiedartifacts))
        MapLatestExtensions = {}
        for k in identifiedartifacts:
            versionResp = self.sendPostRequest(str(k))
            MapLatestExtensions[k] = versionResp
        return MapLatestExtensions

    def sendPostRequest(self, artifact):
        extensionUID = ''
        versionDet = ''
        ARTI = '"' + artifact + '"'
        # print("Tech\Framework checking : " + ARTI)
        determinatorURL = "https://extend.castsoftware.com/V2/services/determinator"
        responseText = {}
        recommendation = {}
        data = {"ids" : [artifact]}
        r = requests.post(determinatorURL, json=data, verify=True)
        # print(r.text)
        splitText = r.text.split(",")
        # print(splitText)
        for a in splitText:
            if a.__contains__("errormessage"):
                return "Unsupported Technology"
            if a.__contains__("recommendedversion"):
                versionDet = a.rsplit(":")
                # print(versionDet)
                versionDet = versionDet[1].replace("\"", "").replace("}", "")
    
            if a.__contains__("extensionuid"):
                extensionUID = a.rsplit(":")
                extensionUID = extensionUID[1].replace("\"", "").replace("}", "")
                # print(extensionUID)
                # print(extensionUID + "--" + versionDet)
        return extensionUID + "--" + versionDet

    def createConnection(self, srcDB, srcHost, srcPort):
    
        # print ("Connecting to database ....  ")
        
        # get a connection, if a connect cannot be made an exception will be raised here
        dbengine = create_postgres_engine(user="operator", host=srcHost, port=srcPort, password="CastAIP")           
        conn = dbengine.connect()
    #     conn = psycopg2.connect(dbname=srcDB, user="operator", host=srcHost,
    #                             port=srcPort, password="CastAIP")
    # 
    #     # conn.cursor will return a cursor object, you can use this cursor to perform queries
        cursor = conn.cursor()
        # print ("Connected!\n")
        return cursor, conn

    def getInstalledExtnsInDB(self, mngtName, srcDB, srcHost, srcPort):
        installedExtenMap = {}
        # cursor, conn = self.createConnection(srcDB, srcHost, srcPort)
        cursor = mngtName.execute_query("select * from sys_package_version")
        for eachrow in cursor:
            installedExtenMap[str(eachrow[0]).replace("/", "")] = eachrow[1]
        return installedExtenMap 
