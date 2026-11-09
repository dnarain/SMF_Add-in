#!/usr/bin/env python3
"""Saved-file and interpretation fixtures; no Excel, network, or third-party modules."""
from __future__ import annotations

from copy import deepcopy
import hashlib
from html import escape
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import zipfile

from read_results import WorkbookReadError, read, redact_anonymous
from make_report import attach_http_diagnostics, build_report, causal_http, interpret, provider_working, version_matches, worksheet_entry, coverage_kind

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE = "http://schemas.openxmlformats.org/package/2006/relationships"


def workbook(path, result_cells, strings="", control_cells="", absolute=False):
    target = "/xl/worksheets/results.xml" if absolute else "worksheets/results.xml"
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("xl/workbook.xml", f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><sheets><sheet name="FunctionResults" sheetId="1" r:id="rId1"/><sheet name="TestControl" sheetId="2" r:id="rId2"/></sheets></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels", f'<Relationships xmlns="{PACKAGE}"><Relationship Id="rId1" Target="{target}"/><Relationship Id="rId2" Target="worksheets/control.xml"/></Relationships>')
        z.writestr("xl/worksheets/results.xml", f'<worksheet xmlns="{MAIN}"><sheetData>{result_cells}</sheetData></worksheet>')
        z.writestr("xl/worksheets/control.xml", f'<worksheet xmlns="{MAIN}"><sheetData>{control_cells}</sheetData></worksheet>')
        z.writestr("xl/vbaProject.bin", b"fixture-vba")
        if strings:
            z.writestr("xl/sharedStrings.xml", f'<sst xmlns="{MAIN}">{strings}</sst>')


def inline(address, value):
    return f'<c r="{address}" t="inlineStr"><is><t>{escape(value)}</t></is></c>'


class ReaderTests(unittest.TestCase):
    def test_rich_and_inline_text_skip_phonetic_guides_and_keep_high_rows(self):
        with tempfile.TemporaryDirectory() as d:
            path = Path(d)/"saved.xlsm"
            cells = '<row r="2501"><c r="A2501" t="s"><v>0</v></c>'+inline("B2501", "FunctionOne")+inline("F2501", "PASS")+'<c r="H2501" t="inlineStr"><is><r><t>A</t></r><rPh><t>PHONETIC</t></rPh><r><t>B</t></r></is></c></row>'
            strings = '<si><r><t>case</t></r><rPh><t>WRONG</t></rPh><r><t>-one</t></r></si>'
            control = '<row r="4">'+inline("A4", "Run ID")+inline("B4", "fixture-run")+'</row>'
            workbook(path, cells, strings, control, absolute=True)
            result = read(path, d)
            self.assertEqual(result["results"][0]["id"], "case-one")
            self.assertEqual(result["results"][0]["actual"], "AB")
            self.assertEqual(result["results"][0]["row"], 2501)
            self.assertEqual(result["run_metadata"]["Run ID"], "fixture-run")
            self.assertEqual(result["workbook_vba_sha256"], hashlib.sha256(b"fixture-vba").hexdigest())

    def test_raw_statuses_and_cached_errors_are_preserved(self):
        statuses = ["PASS", "CHECK_FAILED", "RUNTIME_ERROR", "SETUP_ERROR", "ASSERTION_ERROR", "CLEANUP_ERROR", "PENDING", "RUNNING"]
        with tempfile.TemporaryDirectory() as d:
            path=Path(d)/"saved.xlsm"
            rows="".join(f'<row r="{i}">'+inline(f'A{i}',f'case{i}')+inline(f'F{i}',status)+f'<c r="H{i}" t="e"><v>#VALUE!</v></c></row>' for i,status in enumerate(statuses,2))
            workbook(path,rows)
            result=read(path)
            self.assertEqual([x["raw_status"] for x in result["results"]],statuses)
            self.assertEqual(result["results"][0]["actual"],"#VALUE!")
            self.assertEqual(result["results"][0]["xml_cell_types"]["actual"],"e")

    def test_invalid_shared_string_is_an_explicit_error(self):
        with tempfile.TemporaryDirectory() as d:
            path=Path(d)/"bad.xlsm"
            workbook(path,'<row r="2"><c r="A2" t="s"><v>9</v></c></row>','<si><t>only</t></si>')
            with self.assertRaisesRegex(WorkbookReadError,"shared-string"):
                read(path)

    def test_cli_requires_explicit_inputs_and_runs_outside_repository(self):
        script=Path(__file__).with_name("read_results.py")
        with tempfile.TemporaryDirectory() as d:
            result=subprocess.run([sys.executable,str(script)],cwd=d,capture_output=True,text=True)
            self.assertNotEqual(result.returncode,0)
            self.assertIn("--suite",result.stderr)
            path=Path(d)/"saved.xlsm";workbook(path,'<row r="2">'+inline("A2","case1")+inline("F2","PASS")+'</row>')
            result=subprocess.run([sys.executable,str(script),'--suite',d,'--workbook',str(path)],cwd=d,capture_output=True,text=True)
            self.assertEqual(result.returncode,0,result.stderr)
            private=Path(d)/"native-results.private.json"
            self.assertTrue(private.exists())
            self.assertEqual(private.stat().st_mode & 0o777,0o600)

    def test_report_cli_writes_portable_outputs_and_rejects_wrong_case_hash(self):
        script=Path(__file__).with_name('make_report.py')
        with tempfile.TemporaryDirectory(prefix='suite with spaces ') as d:
            suite=Path(d);path=suite/'saved.xlsm'
            specs=[{'id':'F','function':'F','network':False}]
            for name,value in [('test-cases.json',specs),('function-inventory.json',[{'function':'F','visibility':'Public','file':'code/F.bas'}]),('datasource_cases.json',[]),('datasource_inventory_meta.json',{})]:
                (suite/name).write_text(json.dumps(value))
            row='<row r="2">'+inline('A2','F')+inline('B2','F')+inline('F2','PASS')+'</row>'
            workbook(path,row)
            result=subprocess.run([sys.executable,str(script),'--suite',d,'--workbook',str(path)],cwd='/',capture_output=True,text=True)
            self.assertEqual(result.returncode,0,result.stderr)
            summary=json.loads((suite/'report-summary.json').read_text())
            self.assertEqual(summary['raw_case_counts'],{'PASS':1})
            self.assertFalse(summary['version_verified'])
            self.assertTrue((suite/'endpoint-results.csv').exists())
            self.assertTrue((suite/'uncovered-call-patterns.csv').exists())
            workbook(path,row,control_cells='<row r="11">'+inline('A11','Case-spec SHA')+inline('B11','wrong-hash')+'</row>')
            result=subprocess.run([sys.executable,str(script),'--suite',d,'--workbook',str(path)],cwd='/',capture_output=True,text=True)
            self.assertNotEqual(result.returncode,0)
            self.assertIn('Case-spec SHA',result.stderr)

    def test_actual_runner_hash_label_rejects_edited_specs_and_source_mismatch(self):
        script=Path(__file__).with_name('make_report.py')
        with tempfile.TemporaryDirectory() as d:
            suite=Path(d);path=suite/'saved.xlsm';spec_path=suite/'test-cases.json'
            original=json.dumps([{'id':'F','function':'F','network':False,'expected_text':'original assertion'}])
            spec_path.write_text(original)
            for name,value in [('function-inventory.json',[{'function':'F','visibility':'Public'}]),('datasource_cases.json',[]),('datasource_inventory_meta.json',{})]:
                (suite/name).write_text(json.dumps(value))
            digest=hashlib.sha256(spec_path.read_bytes()).hexdigest();fingerprint='a'*64
            (suite/'run-manifest.json').write_text(json.dumps({'case_specification_sha256':digest,'source_fingerprint':fingerprint}))
            row='<row r="2">'+inline('A2','F')+inline('B2','F')+inline('F2','PASS')+'</row>'
            def write_owner(source):
                control='<row r="10">'+inline('A10','Source fingerprint')+inline('B10',source)+'</row>'
                control+='<row r="11">'+inline('A11','Case specification SHA256')+inline('B11',digest)+'</row>'
                workbook(path,row,control_cells=control)
            def run():
                return subprocess.run([sys.executable,str(script),'--suite',d,'--workbook',str(path)],capture_output=True,text=True)
            write_owner(fingerprint)
            result=run();self.assertEqual(result.returncode,0,result.stderr)
            bindings=json.loads((suite/'report-summary.json').read_text())['input_bindings']
            self.assertTrue(bindings['workbook_case_specification_verified'])
            self.assertTrue(bindings['source_fingerprint_verified'])
            spec_path.write_text(original.replace('original assertion','edited assertion'))
            result=run();self.assertNotEqual(result.returncode,0)
            self.assertIn('Case specification SHA256',result.stderr)
            spec_path.write_text(original);write_owner('b'*64)
            result=run();self.assertNotEqual(result.returncode,0)
            self.assertIn('Source fingerprint',result.stderr)


class InterpretationTests(unittest.TestCase):
    def setUp(self):
        self.case={"id":"live","function":"LiveFunction","group":"live_extraction","domain":"example.test","expression":"LiveFunction()"}
        self.row={"id":"live","function":"LiveFunction","row":2,"status":"CHECK_FAILED","actual":"Error","http_status":"403","http_url":"https://example.test/data","http_reset_marker":"smf-test-case:run-one:1","http_case_id":"live","http_run_id":"run-one","http_observation":"REQUEST"}

    def test_only_matching_case_run_marker_establishes_request(self):
        self.assertTrue(causal_http(self.row,"run-one")[0])
        for field,value in [('http_case_id','old'),('http_run_id','old'),('http_reset_marker','smf-test-case:run-one:8'),('http_observation','NONE')]:
            row=dict(self.row);row[field]=value
            self.assertFalse(causal_http(row,"run-one")[0],field)
            self.assertEqual(interpret(row,self.case,"run-one")[0],"UNCLASSIFIED_FAIL")

    def test_causal_403_is_unavailable_but_200_extraction_error_is_unclassified(self):
        self.assertEqual(interpret(self.row,self.case,"run-one")[0],"PROVIDER_UNAVAILABLE")
        self.row['http_status']='200'
        self.assertEqual(interpret(self.row,self.case,"run-one")[0],"UNCLASSIFIED_FAIL")

    def test_captured_direct_body_marker_is_required_for_200_rejection(self):
        self.row['http_status']='200';self.row['actual']='String: <html>Access Denied</html>'
        self.assertEqual(interpret(self.row,self.case,"run-one")[0],"UNCLASSIFIED_FAIL")
        self.case['test_mode']='live_endpoint_request'
        self.assertEqual(interpret(self.row,self.case,"run-one")[0],"PROVIDER_UNAVAILABLE")

    def test_fixture_failure_is_not_a_provider_outage(self):
        self.case['fixture_source']='local saved response'
        self.assertEqual(interpret(self.row,self.case,"run-one")[0],"CODE_FAILURE")

    def test_obsolete_pass_never_counts_as_a_working_provider(self):
        self.row['status']='PASS';self.row['actual']='String: Obsolete'
        self.case['expected_obsolete']=True
        status,_=interpret(self.row,self.case,"run-one")
        self.assertEqual(status,"OBSOLETE")
        self.row['interpreted_status']=status
        self.assertFalse(provider_working(self.row,self.case))

    def test_pending_and_running_remain_untested(self):
        for status in ['', 'MISSING','PENDING','RUNNING']:
            self.row['status']=status
            self.assertEqual(interpret(self.row,self.case,"run-one")[0],"UNTESTED")
            self.assertFalse(causal_http(self.row,"run-one")[0])

    def test_current_request_to_another_domain_is_not_provider_attribution(self):
        self.row['http_url']='https://another.test/data'
        self.assertTrue(causal_http(self.row,'run-one')[0])
        self.assertEqual(interpret(self.row,self.case,'run-one')[0],'UNCLASSIFIED_FAIL')

    def test_setup_and_assertion_errors_remain_code_failures(self):
        for status in ['SETUP_ERROR','ASSERTION_ERROR','CLEANUP_ERROR']:
            self.row['status']=status
            self.assertEqual(interpret(self.row,self.case,"run-one")[0],"CODE_FAILURE")

    def test_obsolete_text_does_not_hide_a_setup_failure(self):
        self.case['expected_obsolete']=True
        self.row['status']='SETUP_ERROR';self.row['actual']='String: Obsolete'
        self.assertEqual(interpret(self.row,self.case,'run-one')[0],'CODE_FAILURE')

    def test_primary_and_all_case_coverage_are_separate_and_subs_classes_excluded(self):
        specs=[{'id':'F','function':'F','network':False},{'id':'extra','function':'F','network':False},{'id':'G','function':'G','network':False}]
        native={'workbook':'fixture.xlsm','workbook_vba_sha256':'abc','results':[{'id':'F','function':'F','status':'CHECK_FAILED'},{'id':'extra','function':'F','status':'PASS'}],'run_metadata':{'Run ID':'r','Expected SMF version':'3.0.2024.07.30-mac.0.5-dev','Actual SMF version':'Stock Market Functions add-in, Version 3.0.2024.07.30-mac.0.5-dev (/fixture)'}}
        inventory=[{'function':'F','visibility':'Public','file':'code/F.bas'},{'function':'G','visibility':'Public','file':'code/G.bas'},{'function':'SubOne','kind':'sub','visibility':'Public','file':'code/F.bas'},{'function':'Item','visibility':'Public','file':'code/Dictionary.cls'}]
        report=build_report(native,specs,inventory,[],{})
        self.assertEqual(report['summary']['public_functions'],2)
        f=report['functions'][0]
        self.assertEqual(f['primary_raw_status'],'CHECK_FAILED')
        self.assertEqual(f['interpreted_counts'],{'CODE_FAILURE':1,'PASS':1})
        self.assertEqual(report['functions'][1]['primary_interpreted_status'],'UNTESTED')
        self.assertTrue(any(gap['pattern']=='worksheet evaluation' for gap in report['uncovered_call_patterns']))

    def test_named_direct_case_is_primary_but_indirect_check_is_not_promoted(self):
        native={'workbook':'fixture.xlsm','workbook_vba_sha256':'abc','results':[
            {'id':'Repair_Direct','function':'F','status':'PASS'},
            {'id':'IndirectOnly','function':'G','status':'PASS'}],'run_metadata':{}}
        specs=[{'id':'Repair_Direct','function':'F','expression':'F(fixture)','network':False},
               {'id':'IndirectOnly','function':'G','expression':'AnotherFunction()','check':'G(result) = 42','network':False}]
        inventory=[{'function':'F','visibility':'Public'},{'function':'G','visibility':'Public'}]
        report=build_report(native,specs,inventory,[],{})
        self.assertEqual(report['functions'][0]['primary_case_id'],'Repair_Direct')
        self.assertEqual(report['functions'][0]['primary_raw_status'],'PASS')
        self.assertIn('direct-call',report['functions'][0]['primary_selection'])
        self.assertEqual(report['functions'][1]['primary_raw_status'],'MISSING')
        self.assertEqual(report['functions'][1]['all_case_count'],1)

    def test_mismatched_run_identity_does_not_count_saved_passes(self):
        native={'workbook':'fixture.xlsm','workbook_vba_sha256':'abc','results':[{'id':'F','function':'F','status':'PASS'}],'run_metadata':{'Run ID':'r'}}
        report=build_report(native,[{'id':'F','function':'F'}],[{'function':'F','visibility':'Public'}],[],{}, {'run_id':'different'})
        self.assertEqual(report['cases'][0]['raw_status'],'PASS')
        self.assertEqual(report['cases'][0]['interpreted_status'],'UNTESTED')
        self.assertTrue(report['summary']['evidence_errors'])

    def test_version_gate_requires_exact_full_version_before_path(self):
        expected='3.0.2024.07.30-mac.0.6-dev'
        prefix='Stock Market Functions add-in, Version '
        self.assertTrue(version_matches(expected,prefix+expected+' (/fixture)'))
        invalid=[prefix+expected+'-old (/fixture)',
                 prefix+'3.0.2024.07.30-mac.0.5-dev (/path/'+expected+')',
                 prefix+expected,expected+' (/fixture)',
                 'Wrong prefix '+expected+' (/fixture)']
        for actual in invalid:
            self.assertFalse(version_matches(expected,actual),actual)
            native={'workbook':'fixture.xlsm','workbook_vba_sha256':'abc','results':[{'id':'F','function':'F','status':'PASS'}],
                    'run_metadata':{'Expected SMF version':expected,'SMF version':actual}}
            report=build_report(native,[{'id':'F','function':'F'}],[{'function':'F','visibility':'Public'}],[],{})
            self.assertFalse(report['summary']['version_verified'])
            self.assertEqual(report['cases'][0]['interpreted_status'],'UNTESTED')

    def test_35_endpoint_groups_do_not_become_35_working_financial_sources(self):
        specs=[];results=[];sources=[]
        for i in range(1,36):
            case_id=f'endpoint{i}';domain=f'source{i}.test';url=f'https://{domain}/data'
            specs.append({'id':case_id,'function':'Fetch','test_mode':'live_endpoint_request','domain':domain})
            results.append({'id':case_id,'function':'Fetch','row':i+1,'status':'PASS','http_status':'200','http_url':url,'http_reset_marker':f'smf-test-case:r:{i}','http_case_id':case_id,'http_run_id':'r','http_observation':'REQUEST'})
            sources.append({'id':domain,'domain':domain,'probe_urls':[url],'native_cases':[{'id':case_id}]})
        native={'workbook':'fixture.xlsm','workbook_vba_sha256':'abc','results':results,'run_metadata':{'Run ID':'r'}}
        report=build_report(native,specs,[{'function':'Fetch','visibility':'Public'}],sources,{})
        self.assertEqual(report['summary']['source_groups'],35)
        self.assertEqual(report['summary']['positive_live_extraction_passes'],0)
        self.assertTrue(all(source['endpoint_content_passes']==1 and source['positive_extraction_passes']==0 for source in report['sources']))
        self.assertTrue(all(endpoint['coverage_status']=='OBSERVED' for endpoint in report['endpoints']))

    def test_separate_403_probe_does_not_classify_native_failure(self):
        row={'id':'live','function':'LiveFunction','status':'CHECK_FAILED','actual':'Error','http_status':'403','http_url':'https://example.test/data'}
        native={'workbook':'fixture.xlsm','workbook_vba_sha256':'abc','results':[row],'run_metadata':{'Run ID':'r'}}
        sources=[{'id':'example.test','domain':'example.test','native_cases':[{'id':'live'}]}]
        report=build_report(native,[self.case],[{'function':'LiveFunction','visibility':'Public'}],sources,{})
        self.assertEqual(report['cases'][0]['interpreted_status'],'UNCLASSIFIED_FAIL')
        with tempfile.TemporaryDirectory() as d:
            p=Path(d)
            (p/'results.json').write_text(json.dumps([{'index':1,'source_ids':['example.test'],'response':{'http_status':403,'requested_url':'https://example.test/data'}}]))
            (p/'run-metadata.json').write_text(json.dumps({'schema_version':2,'path_base':'this_output_folder'}))
            attach_http_diagnostics(report,p)
        self.assertEqual(report['cases'][0]['interpreted_status'],'UNCLASSIFIED_FAIL')
        self.assertEqual(report['summary']['separate_http_probe_records'],1)
        self.assertEqual(report['sources'][0]['separate_http_probe_counts'],{'403':1})

    def test_observed_causal_alias_host_associates_source_and_exact_endpoint(self):
        url='https://ca.advfn.com/p.php?pid=financials&symbol=NASDAQ:AAPL'
        row={'id':'ADVFN','function':'F','row':2,'status':'CHECK_FAILED','actual':'Error',
             'http_status':'200','http_url':url,'http_reset_marker':'smf-test-case:r:1',
             'http_case_id':'ADVFN','http_run_id':'r','http_observation':'REQUEST'}
        specs=[{'id':'ADVFN','function':'F','group':'live_extraction','domain':'www.advfn.com','expression':'F()'}]
        native={'workbook':'fixture.xlsm','workbook_vba_sha256':'abc','results':[row],'run_metadata':{'Run ID':'r'}}
        sources=[{'domain':'www.advfn.com','probe_urls':['https://www.advfn.com/']},
                 {'domain':'ca.advfn.com','probe_urls':[url]},
                 {'domain':'unrecorded-cdn.test','probe_urls':['https://unrecorded-cdn.test/image.png']}]
        report=build_report(native,specs,[{'function':'F','visibility':'Public'}],sources,{})
        self.assertEqual(report['cases'][0]['raw_status'],'CHECK_FAILED')
        self.assertEqual(report['cases'][0]['interpreted_status'],'UNCLASSIFIED_FAIL')
        groups={source['domain']:source for source in report['sources']}
        self.assertEqual(groups['www.advfn.com']['case_ids'],['ADVFN'])
        self.assertEqual(groups['ca.advfn.com']['case_ids'],['ADVFN'])
        self.assertEqual(groups['unrecorded-cdn.test']['case_ids'],[])
        endpoint=next(item for item in report['endpoints'] if item['source']=='ca.advfn.com')
        self.assertEqual(endpoint['case_ids'],['ADVFN'])
        self.assertEqual(endpoint['coverage_status'],'OBSERVED')
        # A stale same-host URL cannot create the new observed-host association.
        row['http_observation']='NONE'
        stale=build_report(native,specs,[{'function':'F','visibility':'Public'}],sources,{})
        self.assertEqual(next(item for item in stale['sources'] if item['domain']=='ca.advfn.com')['case_ids'],[])

    def test_worksheet_entry_is_independent_of_fixture_and_vba_call_categories(self):
        cases=[{'id':'smfInCellChart','setup':['chartCell.FormulaArray = formulaText']},
               {'id':'RCHCreateComment','setup':['commentCell.Formula = formulaText']},
               {'id':'MacComment_failed_update_clears_picture','fixture_source':'local PNG','call_steps':['target.Formula2 = formulaText']},
               {'id':'ImageComment_loopback','domain':'loopback','setup':['target.Formula2 = formulaText']}]
        self.assertEqual([coverage_kind(case) for case in cases],['VBA_CALL','VBA_CALL','FIXTURE','FIXTURE'])
        self.assertTrue(all(worksheet_entry(case) for case in cases))
        self.assertTrue(worksheet_entry({'worksheet_formula':'=F()'}))
        self.assertTrue(worksheet_entry({'setup':['If ready Then target.Formula2 = formulaText']}))
        for case in [{'worksheet_formula':None},
                     {'setup':['oldFormula = target.Formula2']},
                     {'cleanup':['target.Formula2 = oldFormula']},
                     {'setup':['If target.Formula2 = "old" Then result = True']},
                     {'setup':['comparison = target.Formula2 = "old"']},
                     {'setup':['Inspect(target.Formula2 = "old")']},
                     {'setup':['note = "target.Formula2 = value"']}]:
            self.assertFalse(worksheet_entry(case),case)

    def test_404_reason_does_not_claim_a_provider_wide_outage(self):
        self.row['http_status']='404'
        status,reason=interpret(self.row,self.case,'run-one')
        self.assertEqual(status,'PROVIDER_UNAVAILABLE')
        self.assertIn('requested URL is correct',reason)
        self.assertIn('does not establish',reason)

    def test_malformed_url_is_code_failure_only_with_case_bound_request(self):
        self.row['http_status']='0'
        self.row['http_error']='Request failed (3): curl: (3) URL rejected: Malformed input to a URL function\n'
        status,reason=interpret(self.row,self.case,'run-one')
        self.assertEqual(status,'CODE_FAILURE')
        self.assertIn('supplied or constructed an invalid request URL',reason)
        self.row['http_observation']='NONE'
        self.assertEqual(interpret(self.row,self.case,'run-one')[0],'UNCLASSIFIED_FAIL')
        self.row['http_observation']='REQUEST';self.row['http_status']='200'
        self.assertEqual(interpret(self.row,self.case,'run-one')[0],'UNCLASSIFIED_FAIL')


class RedactionTests(unittest.TestCase):
    def test_nested_session_values_are_removed_but_statuses_remain(self):
        value={'url':'https://query1.finance.yahoo.com/data?x=1&crumb=CRUMB_SENTINEL&y=2',
               'error':'Cookie: A1=COOKIE_SENTINEL; A3=OTHER_SENTINEL\nrequest failed',
               'nested':[{'crumb':'DIRECT_SENTINEL','cookies':{'A1':'MAP_SENTINEL'},'cookie_received':True,'crumb_status':200}],
               'bare':'A1=d=COOKIE_SENTINEL&a=TAIL_SENTINEL; crumb="BARE_SENTINEL"',
               'request':{'requested_url':'https://query1.finance.yahoo.com/v1/test/getcrumb','actual':'UNLABELED_SENTINEL'}}
        redacted=redact_anonymous(value);serialized=json.dumps(redacted)
        self.assertNotIn('SENTINEL',serialized)
        self.assertTrue(redacted['nested'][0]['cookie_received'])
        self.assertEqual(redacted['nested'][0]['crumb_status'],200)

    def test_report_outputs_are_redacted_and_raw_file_is_private(self):
        script=Path(__file__).with_name('make_report.py')
        with tempfile.TemporaryDirectory() as d:
            suite=Path(d);path=suite/'saved.xlsm'
            for name,value in [('test-cases.json',[{'id':'F','function':'F','network':False}]),('function-inventory.json',[{'function':'F','visibility':'Public'}]),('datasource_cases.json',[]),('datasource_inventory_meta.json',{})]:
                (suite/name).write_text(json.dumps(value))
            actual='Cookie: A1=COOKIE_SENTINEL;'
            row='<row r="2">'+inline('A2','F')+inline('B2','F')+inline('F2','CHECK_FAILED')+inline('H2',actual)+inline('K2','https://query1.finance.yahoo.com/data?crumb=CRUMB_SENTINEL')+'</row>'
            workbook(path,row)
            http=suite/'http';http.mkdir()
            (http/'results.json').write_text(json.dumps([{'index':1,'source_ids':['source.test'],'response':{'http_status':403,'requested_url':'https://query1.finance.yahoo.com/data?crumb=HTTP_SENTINEL','error':'Cookie: A1=HTTP_COOKIE_SENTINEL;'}}]))
            (http/'run-metadata.json').write_text(json.dumps({'schema_version':2}))
            result=subprocess.run([sys.executable,str(script),'--suite',d,'--workbook',str(path),'--http-dir',str(http)],capture_output=True,text=True)
            self.assertEqual(result.returncode,0,result.stderr)
            private=suite/'native-results.private.json'
            self.assertIn('COOKIE_SENTINEL',private.read_text())
            self.assertEqual(private.stat().st_mode & 0o777,0o600)
            for name in ['native-results.json','report-data.json','report-summary.json','test-results.md','all-case-results.csv','http-probe-results.csv']:
                self.assertNotIn('SENTINEL',(suite/name).read_text(),name)


if __name__ == '__main__':
    unittest.main()
