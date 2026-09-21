"""BLUE CAB panel v1.8. No emails or Drive files are written by this application."""
from __future__ import annotations
import base64, io, os, secrets, shutil, tempfile, zipfile
from pathlib import Path
from flask import Flask, request, render_template, redirect, url_for, session, flash, send_file, abort
from werkzeug.utils import secure_filename
from bluecab_engine import process_files, normalized_date, APP_VERSION, DEFAULT_MASTER, TZ, MasterError
from list_duzenle import normalize_workbook
import master_store

BASE=Path(__file__).resolve().parent
PANEL_VERSION=APP_VERSION
GMAIL_SCOPE='https://www.googleapis.com/auth/gmail.readonly'
SHEETS_SCOPE='https://www.googleapis.com/auth/spreadsheets.readonly'
MASTER_SHEET_ID=os.environ.get('BLUECAB_MASTER_SHEET_ID','1xMpQE4U6EIS3e3N-NkCni3EWdKANJ_FV06T-2c685iQ')
REDIRECT_URI=os.environ.get('GOOGLE_REDIRECT_URI','https://blue-cab-web-panel.onrender.com/oauth2callback')
app=Flask(__name__)
app.secret_key=os.environ.get('FLASK_SECRET_KEY') or secrets.token_hex(32)
app.config.update(MAX_CONTENT_LENGTH=100*1024*1024,SESSION_COOKIE_HTTPONLY=True,
                  SESSION_COOKIE_SAMESITE='Lax',SESSION_COOKIE_SECURE=os.environ.get('RENDER')=='true')
# Preserve the original single-process, per-session credential behavior. No token in browser cookies.
OAUTH_STORE={}


def _sid():
    if 'sid' not in session:session['sid']=secrets.token_urlsafe(24)
    return session['sid']


def csrf_token():
    if 'csrf' not in session:session['csrf']=secrets.token_urlsafe(32)
    return session['csrf']


app.jinja_env.globals['csrf_token']=csrf_token


@app.before_request
def protect_post():
    if request.method=='POST':
        expected=session.get('csrf','')
        actual=request.form.get('csrf_token','')
        if not expected or not secrets.compare_digest(expected,actual):
            abort(400,description='Oturum kontrol\u00fc ge\u00e7ersiz. Paneli yenileyip tekrar deneyin.')


@app.after_request
def safe_headers(response):
    response.headers['X-Content-Type-Options']='nosniff'
    response.headers['X-Frame-Options']='SAMEORIGIN'
    response.headers['Referrer-Policy']='same-origin'
    if not request.path.startswith('/static'):response.headers['Cache-Control']='no-store'
    return response


def _admin_enabled():
    return len(os.environ.get('BLUECAB_ADMIN_KEY',''))>=16


def _require_admin():
    key=os.environ.get('BLUECAB_ADMIN_KEY','')
    supplied=request.form.get('admin_key','')
    if len(key)<16:abort(403,description='MASTER de\u011fi\u015fikli\u011fi kilitli. BLUECAB_ADMIN_KEY en az 16 karakter olmal\u0131.')
    if not secrets.compare_digest(key,supplied):abort(403,description='Y\u00f6netici anahtar\u0131 do\u011frulanamad\u0131.')


def _client_config():
    cid=os.environ.get('GOOGLE_CLIENT_ID'); secret=os.environ.get('GOOGLE_CLIENT_SECRET')
    if not cid or not secret:raise ValueError('GOOGLE_CLIENT_ID / GOOGLE_CLIENT_SECRET ayarlar\u0131 eksik.')
    return {'web':{'client_id':cid,'client_secret':secret,'auth_uri':'https://accounts.google.com/o/oauth2/auth',
                   'token_uri':'https://oauth2.googleapis.com/token','redirect_uris':[REDIRECT_URI]}}


def _store_creds(creds):
    OAUTH_STORE[_sid()]={'token':creds.token,'refresh_token':creds.refresh_token,
                        'token_uri':creds.token_uri,'client_id':creds.client_id,
                        'client_secret':creds.client_secret,'scopes':list(creds.scopes or [])}


def _creds():
    data=OAUTH_STORE.get(_sid())
    if not data:return None
    from google.oauth2.credentials import Credentials
    from google.auth.transport.requests import Request
    creds=Credentials(**data)
    if creds.expired and creds.refresh_token:creds.refresh(Request()); _store_creds(creds)
    return creds


def _gmail_service():
    from googleapiclient.discovery import build
    creds=_creds()
    if not creds or not creds.has_scopes([GMAIL_SCOPE]):return None
    return build('gmail','v1',credentials=creds,cache_discovery=False)


@app.get('/')
def index():
    info={}; error=''
    try:_,info=master_store.active()
    except Exception as exc:error=str(exc)
    data=OAUTH_STORE.get(_sid(),{})
    return render_template('index.html',version=PANEL_VERSION,master=info,master_error=error,
                           gmail_connected=GMAIL_SCOPE in data.get('scopes',[]),
                           sheets_connected=SHEETS_SCOPE in data.get('scopes',[]),
                           admin_enabled=_admin_enabled(),persistent_configured=bool(os.environ.get('BLUECAB_DATA_DIR')))


def _uploads(work):
    paths=[]
    for i,f in enumerate(request.files.getlist('files'),1):
        if not f or not f.filename:continue
        if not f.filename.lower().endswith(('.xlsx','.xlsm')):raise ValueError('Yaln\u0131zca .xlsx/.xlsm dosyalar\u0131 i\u015flenebilir.')
        folder=work/'input'/str(i); folder.mkdir(parents=True,exist_ok=True)
        path=folder/(secure_filename(f.filename) or f'liste_{i}.xlsx')
        f.save(path); paths.append(path)
    if not paths:raise ValueError('En az bir Excel dosyas\u0131 se\u00e7in.')
    return paths


def _response_file(path,name=None):
    data=io.BytesIO(Path(path).read_bytes())
    mime='application/zip' if Path(path).suffix=='.zip' else 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    return send_file(data,as_attachment=True,download_name=name or Path(path).name,mimetype=mime)


def _run_action(action):
    with tempfile.TemporaryDirectory(prefix='bluecab_') as td:
        work=Path(td); out=work/'output'; out.mkdir()
        try:
            master,_=master_store.active()  # One immutable snapshot for the WHOLE request.
            selected=request.form.get('master_version','')
            if selected!=master.version:raise MasterError('MASTER de\u011fi\u015fti. Sayfay\u0131 yenileyin; farkl\u0131 s\u00fcr\u00fcmle atama ba\u015flat\u0131lmad\u0131.')
            paths=_uploads(work)
            if action=='process':
                z,_,_=process_files(paths,out,master,normalized_date(request.form.get('date') or None))
                return _response_file(z)
            outputs=[]
            for i,p in enumerate(paths,1):
                dest=out/f'{i:02d}_{p.stem}_DUZENLENMIS.xlsx'
                normalize_workbook(p,dest,master); outputs.append(dest)
            if len(outputs)==1:return _response_file(outputs[0])
            z=work/'BLUE_CAB_DUZENLENMIS_LISTELER.zip'
            with zipfile.ZipFile(z,'w',zipfile.ZIP_DEFLATED) as archive:
                for p in outputs:archive.write(p,p.name)
            return _response_file(z)
        except Exception as exc:
            app.logger.warning('List processing stopped: %s',type(exc).__name__)
            flash('Atama durduruldu: '+str(exc))
            return redirect(url_for('index'))


@app.post('/process')
def process():return _run_action('process')


@app.post('/organize')
def organize():return _run_action('organize')


@app.post('/master/upload')
def master_upload():
    _require_admin()
    uploaded=request.files.get('master_file')
    if not uploaded or not uploaded.filename or not uploaded.filename.lower().endswith('.xlsx'):
        flash('MASTER i\u00e7in tam .xlsx dosyas\u0131n\u0131 se\u00e7in.'); return redirect(url_for('index'))
    with tempfile.TemporaryDirectory(prefix='bluecab_master_') as td:
        p=Path(td)/'MASTER.xlsx'; uploaded.save(p)
        try:
            info=master_store.activate(p,'Excel y\u00fckleme: '+secure_filename(uploaded.filename),request.form.get('master_version',''))
            flash(f'MASTER etkinle\u015ftirildi: {info["rows"]} mahalle, s\u00fcr\u00fcm {info["version"]}.')
        except Exception as exc:flash('MASTER de\u011fi\u015ftirilmedi: '+str(exc))
    return redirect(url_for('index'))


def _oauth_start(scopes,destination):
    from google_auth_oauthlib.flow import Flow
    flow=Flow.from_client_config(_client_config(),scopes=scopes,redirect_uri=REDIRECT_URI)
    auth_url,state=flow.authorization_url(access_type='offline',include_granted_scopes='true',prompt='consent')
    session.update(oauth_state=state,oauth_scopes=scopes,oauth_destination=destination)
    # google-auth-oauthlib versions enabling PKCE require the same verifier at callback.
    session['oauth_verifier']=getattr(flow,'code_verifier',None)
    return redirect(auth_url)


@app.get('/gmail/connect')
def gmail_connect():
    try:return _oauth_start([GMAIL_SCOPE],'gmail_messages')
    except Exception as exc:flash('Google ba\u011flant\u0131s\u0131: '+str(exc)); return redirect(url_for('index'))


@app.post('/master/connect')
def master_connect():
    _require_admin()
    try:return _oauth_start([GMAIL_SCOPE,SHEETS_SCOPE],'index')
    except Exception as exc:flash('Google ba\u011flant\u0131s\u0131: '+str(exc)); return redirect(url_for('index'))


@app.get('/oauth2callback')
def oauth2callback():
    try:
        expected=session.pop('oauth_state',None)
        actual=request.args.get('state','')
        if not expected or not secrets.compare_digest(expected,actual):raise ValueError('OAuth state ge\u00e7ersiz.')
        from google_auth_oauthlib.flow import Flow
        scopes=session.pop('oauth_scopes',[GMAIL_SCOPE])
        verifier=session.pop('oauth_verifier',None)
        flow=Flow.from_client_config(_client_config(),scopes=scopes,state=expected,redirect_uri=REDIRECT_URI,code_verifier=verifier)
        # Use the configured callback origin, not proxy-dependent http request.url.
        flow.fetch_token(authorization_response=REDIRECT_URI+'?'+request.query_string.decode('ascii'))
        _store_creds(flow.credentials); flash('Google salt okunur ba\u011flant\u0131 tamamland\u0131.')
        dest=session.pop('oauth_destination','index')
        return redirect(url_for(dest if dest in ('index','gmail_messages') else 'index'))
    except Exception as exc:flash('Google ba\u011flant\u0131s\u0131 tamamlanamad\u0131: '+str(exc)); return redirect(url_for('index'))


@app.post('/gmail/disconnect')
def gmail_disconnect():
    OAUTH_STORE.pop(_sid(),None); flash('Google ba\u011flant\u0131s\u0131 kapat\u0131ld\u0131.'); return redirect(url_for('index'))


def _walk_parts(payload):
    yield payload
    for part in payload.get('parts',[]) or []:yield from _walk_parts(part)


def _excel_attachments(service,message_id,work):
    msg=service.users().messages().get(userId='me',id=message_id,format='full').execute()
    paths=[]
    for i,part in enumerate(_walk_parts(msg.get('payload',{})),1):
        name=part.get('filename','')
        if not name.lower().endswith(('.xlsx','.xlsm')):continue
        body=part.get('body',{}) or {}; encoded=body.get('data')
        if not encoded and body.get('attachmentId'):
            encoded=service.users().messages().attachments().get(userId='me',messageId=message_id,id=body['attachmentId']).execute().get('data')
        if not encoded:continue
        raw=base64.urlsafe_b64decode(encoded+'='*((-len(encoded))%4))
        if len(raw)>100*1024*1024:raise ValueError('E-posta eki \u00e7ok b\u00fcy\u00fck.')
        folder=work/str(i); folder.mkdir()
        p=folder/(secure_filename(name) or f'ek_{i}.xlsx'); p.write_bytes(raw); paths.append(p)
    return paths


@app.get('/gmail/messages')
def gmail_messages():
    try:
        service=_gmail_service()
        if not service:return redirect(url_for('gmail_connect'))
        response=service.users().messages().list(userId='me',q='has:attachment newer_than:14d',maxResults=30).execute()
        rows=[]
        for item in response.get('messages',[]):
            msg=service.users().messages().get(userId='me',id=item['id'],format='metadata',metadataHeaders=['Subject','From','Date']).execute()
            h={x['name'].lower():x['value'] for x in msg.get('payload',{}).get('headers',[])}
            rows.append({'id':item['id'],'subject':h.get('subject','(konu yok)'),'from':h.get('from',''),'date':h.get('date','')})
        m,_=master_store.active()
        return render_template('gmail_messages.html',rows=rows,master_version=m.version)
    except Exception as exc:flash('E-postalar okunamad\u0131: '+str(exc)); return redirect(url_for('index'))


@app.post('/gmail/process/<message_id>')
def gmail_process(message_id):
    with tempfile.TemporaryDirectory(prefix='bluecab_gmail_') as td:
        try:
            service=_gmail_service()
            if not service:return redirect(url_for('gmail_connect'))
            master,_=master_store.active()
            if request.form.get('master_version')!=master.version:raise MasterError('MASTER de\u011fi\u015fti; sayfay\u0131 yenileyin.')
            work=Path(td); paths=_excel_attachments(service,message_id,work)
            if not paths:raise ValueError('Bu e-postada i\u015flenebilir Excel eki yok.')
            z,_,_=process_files(paths,work/'output',master,normalized_date(request.form.get('date') or None))
            return _response_file(z)
        except Exception as exc:flash('Gmail i\u015flemi durduruldu: '+str(exc)); return redirect(url_for('index'))


@app.post('/master/refresh')
def master_refresh():
    _require_admin()
    with tempfile.TemporaryDirectory(prefix='bluecab_sync_') as td:
        try:
            creds=_creds()
            if not creds or not creds.has_scopes([SHEETS_SCOPE]):raise ValueError('\u00d6nce Drive MASTER okuma iznini ba\u011flay\u0131n.')
            from googleapiclient.discovery import build
            from openpyxl import Workbook
            service=build('sheets','v4',credentials=creds,cache_discovery=False)
            meta=service.spreadsheets().get(spreadsheetId=MASTER_SHEET_ID,fields='properties(title),sheets(properties)').execute()
            target='\u0130stanbul Mahalle MASTER'
            tab=next((x['properties'] for x in meta.get('sheets',[]) if x['properties']['title']==target),None)
            if not tab:raise ValueError('Drive dosyas\u0131nda \u0130stanbul Mahalle MASTER sayfas\u0131 bulunamad\u0131.')
            bound=tab.get('gridProperties',{}).get('rowCount',0)
            if not 1<=bound<=20000:raise ValueError('MASTER sayfa boyutu do\u011frulanamad\u0131.')
            values=service.spreadsheets().values().get(spreadsheetId=MASTER_SHEET_ID,range=f"'{target}'!A1:K{bound}",valueRenderOption='UNFORMATTED_VALUE').execute().get('values',[])
            wb=Workbook(); ws=wb.active; ws.title=target
            for row in values:ws.append(row)
            path=Path(td)/'MASTER.xlsx'; wb.save(path)
            info=master_store.activate(path,'Drive Sheets: '+MASTER_SHEET_ID,request.form.get('master_version',''))
            flash(f'Drive MASTER yenilendi: {info["rows"]} mahalle, s\u00fcr\u00fcm {info["version"]}.')
        except Exception as exc:flash('Drive MASTER etkinle\u015ftirilmedi; mevcut s\u00fcr\u00fcm korundu: '+str(exc))
    return redirect(url_for('index'))


@app.get('/health')
def health():
    try:
        m,_=master_store.active()
        return {'status':'ok','panel':PANEL_VERSION,'master':m.summary()}
    except Exception:
        return {'status':'error','panel':PANEL_VERSION,'error':'MASTER invalid; assignment disabled'},503


if __name__=='__main__':
    app.run(host='127.0.0.1',port=int(os.environ.get('PORT','8080')),debug=False)
