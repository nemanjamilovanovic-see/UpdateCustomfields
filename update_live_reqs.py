from pathlib import Path
import sys
import requests
from requests.adapters import HTTPAdapter
import yaml
import tkinter as tk
from tkinter import filedialog
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import logging
import time
import urllib3
from urllib3.exceptions import InsecureRequestWarning

# Suppress insecure TLS warnings when verify=False
urllib3.disable_warnings(InsecureRequestWarning)

logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

def choose_excel_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Izaberite Excel fajl",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    return file_path


def get_access_token(conf, username, pwd):
    request_body_token = {'loginName': username,
                          'password': pwd }
    # Do not log passwords; only log the username at DEBUG level
    logger.debug("Login attempt for user=%s", username)
    url = conf['live_url'] + 'login'
    logger.debug("Auth URL: %s", url)
    response_token = requests.post(url, json=request_body_token, headers={
        'Content-Type': 'application/json',
        'cache-control': 'no-cache'
    }, verify=False)
    if response_token.status_code == 200:
        at = response_token.json().get('accessToken')
        # Do not log tokens; just confirm success
        logger.info('Authentication successful')
        return at
    else:
        logger.error('Authentication failed: code=%s text=%s', response_token.status_code, response_token.text)

def update_live_requests(conf, username, pwd,  liverequests, progress_callback=None):
    start_ts = time.perf_counter()
    # ssl verification paths (kept disabled; remove noisy print)
    url = conf['live_url']
    t0 = time.perf_counter()
    access_token = get_access_token (conf, username, pwd);
    login_s = time.perf_counter() - t0

    def _coerce_id(val):
        if isinstance(val, dict):
            return val.get('id')
        return str(val) if val is not None else None

    def _ensure_list_of_id_objs(cat):
        out = []
        if isinstance(cat, list):
            for c in cat:
                cid = _coerce_id(c)
                if cid:
                    out.append({"id": str(cid)})
        elif cat is not None:
            cid = _coerce_id(cat)
            if cid:
                out.append({"id": str(cid)})
        return out

    services_url = url + 'services'
    # Separate timeouts (configurable via YAML): connect, read
    connect_timeout = float(conf.get('connect_timeout', 10)) if isinstance(conf, dict) else 10.0
    search_read_timeout = float(conf.get('search_read_timeout', 30)) if isinstance(conf, dict) else 30.0
    save_read_timeout = float(conf.get('save_read_timeout', 60)) if isinstance(conf, dict) else 60.0
    search_timeout = (connect_timeout, search_read_timeout)
    save_timeout = (connect_timeout, save_read_timeout)
    # Align connection pool with configured concurrency
    pool_size = int(conf.get('concurrency', 12)) if isinstance(conf, dict) else 12
    # Configurable retry backoff for timeouts (seconds)
    backoff_s = float(conf.get('retry_backoff_s', 0.3)) if isinstance(conf, dict) else 0.3

    # Shared performance stats (thread-safe)
    stats = {
        'login_s': login_s,
        'search_count': 0,
        'search_s_total': 0.0,
        'search_s_min': float('inf'),
        'search_s_max': 0.0,
        'search_fail': 0,
        'save_count': 0,
        'save_s_total': 0.0,
        'save_s_min': float('inf'),
        'save_s_max': 0.0,
        'save_fail': 0,
        'save_timeouts': 0,
        'save_timeout_retries': 0,
        'save_conflict_retries': 0,
        'skipped_unchanged': 0,
        'vendor_fallback': 0,
    }
    stats_lock = threading.Lock()

    def make_session():
        s = requests.Session()
        adapter = HTTPAdapter(pool_connections=pool_size, pool_maxsize=pool_size, max_retries=0)
        s.mount('https://', adapter)
        s.mount('http://', adapter)
        s.verify = False
        s.headers.update({
            'Content-Type': 'application/json',
            'Accept': 'application/json',
            'Authorization': f'Bearer {access_token}'
        })
        return s

    # Reuse a single Session per thread for connection pooling efficiency
    _thread_local = threading.local()

    def get_session():
        s = getattr(_thread_local, 'session', None)
        if s is None:
            s = make_session()
            _thread_local.session = s
        return s

    def process_one(r_str, godina_str):
        session = get_session()
        logger.debug("Processing ID=%s year=%s", r_str, godina_str)

        # 1) Dohvati kompletan zapis preko CPM.searchTopics
        search_body = {
            "application": "CPM",
            "method": "searchTopics",
            "parameters": {
                "filter": [
                    {"field": "id", "operator": "equals", "value": r_str}
                ],
                "paging": []
            }
        }
        t_search0 = time.perf_counter()
        resp_search = session.post(services_url, json=search_body, timeout=search_timeout)
        t_search = time.perf_counter() - t_search0
        with stats_lock:
            stats['search_count'] += 1
            stats['search_s_total'] += t_search
            stats['search_s_min'] = min(stats['search_s_min'], t_search)
            stats['search_s_max'] = max(stats['search_s_max'], t_search)
        try:
            search_json = resp_search.json()
        except Exception:
            search_json = None

        if resp_search.status_code != 200 or not search_json:
            logger.warning("Search failed for ID=%s (status=%s)", r_str, resp_search.status_code)
            with stats_lock:
                stats['search_fail'] += 1
            return False

        results = search_json.get('result') or search_json.get('results') or []
        if not results:
            logger.info("Skip ID=%s: not found in search", r_str)
            return False

        rec = results[0]

        # Ako je godina već ista, preskačemo snimanje radi brzine
        current_year = None
        try:
            current_year = (rec.get('customfields') or {}).get('YearInMaintenance')
        except Exception:
            current_year = None
        if current_year is not None and str(current_year).strip() == godina_str:
            logger.info("Skip ID=%s: already set to %s", r_str, godina_str)
            with stats_lock:
                stats['skipped_unchanged'] += 1
            return True

        # 2) Izgradi kompletan recordData iz rezultata + promeni YearInMaintenance
        name = rec.get('name')
        topic_type = _coerce_id(rec.get('topic_type')) or rec.get('topic_type')
        state = _coerce_id(rec.get('state')) or rec.get('state')
        category = _ensure_list_of_id_objs(rec.get('category'))
        addressbook = _coerce_id(rec.get('addressbook')) or rec.get('addressbook')
        company = _coerce_id(rec.get('company')) or rec.get('company')
        responsible = _coerce_id(rec.get('responsible')) or rec.get('responsible')
        impact = _coerce_id(rec.get('impact')) or rec.get('impact')
        urgency = _coerce_id(rec.get('urgency')) or rec.get('urgency')
        priority = _coerce_id(rec.get('priority')) or rec.get('priority')
        # vendor: pokušaj iz zapisa, u suprotnom fallback na addressbook id (FK ka addressbook.id)
        vendor = _coerce_id(rec.get('vendor')) or _coerce_id(rec.get('addressbook'))
        if not _coerce_id(rec.get('vendor')) and _coerce_id(rec.get('addressbook')):
            with stats_lock:
                stats['vendor_fallback'] += 1

        missing = []
        for fld, val in [("name", name), ("topic_type", topic_type), ("state", state), ("category", category), ("addressbook", addressbook), ("company", company), ("responsible", responsible), ("impact", impact), ("urgency", urgency), ("priority", priority)]:
            if not val:
                missing.append(fld)
        if missing:
            logger.warning("Skip ID=%s: missing mandatory fields: %s", r_str, ', '.join(missing))
            return False

        # pripremi customfields: zadrži postojeće, setuj YearInMaintenance
        cf = rec.get('customfields') or {}
        if not isinstance(cf, dict):
            cf = {}
        cf['YearInMaintenance'] = godina_str

        record_data = {
            "id": r_str,
            "name": name,
            "topic_type": int(topic_type) if str(topic_type).isdigit() else topic_type,
            "state": int(state) if str(state).isdigit() else state,
            "category": category,
            "addressbook": str(addressbook),
            "company": _coerce_id(company) or str(company),
            "responsible": _coerce_id(responsible) or str(responsible),
            "impact": _coerce_id(impact) or str(impact),
            "urgency": _coerce_id(urgency) or str(urgency),
            "priority": _coerce_id(priority) or str(priority),
            "customfields": cf
        }
        if vendor:
            record_data["vendor"] = str(vendor)

        # seq token for optimistic locking if required by backend
        seq_val = rec.get('seq')
        if seq_val is not None:
            record_data['seq'] = int(seq_val) if str(seq_val).isdigit() else seq_val

        def _post_save(body, timeout):
            resp = session.post(services_url, json=body, timeout=timeout)
            try:
                j = resp.json()
            except Exception:
                j = None
            return resp, j

        def _is_concurrency(resp, j):
            if j and str(j.get('errorCode')) == '409':
                return True
            try:
                txt = (resp.text or '').lower()
            except Exception:
                txt = ''
            return 'concurrency conflict' in txt

        def _is_success(resp, j):
            if resp.status_code != 200:
                return False
            if not j:
                return True
            ec = j.get('errorCode') if isinstance(j, dict) else None
            if ec is None:
                return True
            try:
                return int(ec) == 0
            except Exception:
                # treat falsy values (None, 0, '0', '') as success
                return not bool(ec)

        # Primary attempt (services, object parameters)
        primary_body = {
            "application": "CPM",
            "method": "saveTopic",
            "parameters": {
                "recordData": record_data,
                "ignoreLock": True
            }
        }
        t_save_total = 0.0
        try:
            t_save0 = time.perf_counter()
            resp_save, j_save = _post_save(primary_body, save_timeout)
            t_save_total += (time.perf_counter() - t_save0)
        except requests.exceptions.ReadTimeout:
            with stats_lock:
                stats['save_timeouts'] += 1
            logger.warning("Save timeout for ID=%s; verifying state then retrying once if needed", r_str)
            # Verify if value was actually saved despite timeout
            try:
                resp_verify = session.post(services_url, json=search_body, timeout=search_timeout)
                rec_v = None
                if resp_verify.status_code == 200:
                    jv = resp_verify.json()
                    res_v = jv.get('result') or jv.get('results') or []
                    if res_v:
                        rec_v = res_v[0]
                if rec_v:
                    cur = (rec_v.get('customfields') or {}).get('YearInMaintenance')
                    if str(cur).strip() == godina_str:
                        logger.info("Timeout but value persisted for ID=%s; treating as success", r_str)
                        with stats_lock:
                            stats['save_count'] += 1
                        return True
            except Exception:
                pass
            # Retry once after small backoff (configurable)
            if backoff_s > 0:
                time.sleep(backoff_s)
            with stats_lock:
                stats['save_timeout_retries'] += 1
            try:
                t_retry0 = time.perf_counter()
                resp_save, j_save = _post_save(primary_body, save_timeout)
                t_save_total += (time.perf_counter() - t_retry0)
            except requests.exceptions.ReadTimeout:
                logger.error("Save timed out twice for ID=%s", r_str)
                with stats_lock:
                    stats['save_fail'] += 1
                return False

        # If concurrency conflict, refresh seq and retry once
        if _is_concurrency(resp_save, j_save):
            # Refresh search to get latest seq
            resp_search2 = session.post(services_url, json=search_body, timeout=search_timeout)
            try:
                search_json2 = resp_search2.json()
            except Exception:
                search_json2 = None
            if resp_search2.status_code == 200 and search_json2:
                results2 = search_json2.get('result') or search_json2.get('results') or []
                if results2:
                    rec2 = results2[0]
                    seq2 = rec2.get('seq')
                    if seq2 is not None:
                        record_data['seq'] = int(seq2) if str(seq2).isdigit() else seq2
                    # Retry primary with refreshed seq
                    with stats_lock:
                        stats['save_conflict_retries'] += 1
                    t_retry0 = time.perf_counter()
                    resp_save, j_save = _post_save(primary_body, save_timeout)
                    t_save_total += (time.perf_counter() - t_retry0)

        with stats_lock:
            stats['save_count'] += 1
            stats['save_s_total'] += t_save_total
            stats['save_s_min'] = min(stats['save_s_min'], t_save_total)
            stats['save_s_max'] = max(stats['save_s_max'], t_save_total)

        if _is_success(resp_save, j_save):
            logger.info("Updated ID=%s → %s", r_str, godina_str)
            return True
        else:
            err = None
            if j_save and j_save.get('errorMessage'):
                err = j_save.get('errorMessage')
            logger.error("Update failed for ID=%s (status=%s) %s", r_str, resp_save.status_code, err or '')
            with stats_lock:
                stats['save_fail'] += 1
            return False

    # Parallel execution
    max_workers = int(conf.get('concurrency', 10)) if isinstance(conf, dict) else 10
    futures = []
    future_map = {}
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        for r, godina in liverequests.items():
            r_str = str(r).strip()
            godina_str = str(godina).strip()
            fut = executor.submit(process_one, r_str, godina_str)
            futures.append(fut)
            future_map[fut] = (r_str, godina_str)

        # Wait for all to complete (optional aggregation)
        success = 0
        total = len(futures)
        for f in as_completed(futures):
            try:
                res = f.result()
                if progress_callback:
                    rid, yr = future_map.get(f, (None, None))
                    try:
                        progress_callback(rid, yr, bool(res))
                    except Exception:
                        pass
                if res:
                    success += 1
            except Exception as e:
                logger.exception("Worker thread error: %s", e)
                if progress_callback:
                    rid, yr = future_map.get(f, (None, None))
                    try:
                        progress_callback(rid, yr, False)
                    except Exception:
                        pass
        total_elapsed = time.perf_counter() - start_ts
        logger.info("Finished: %s/%s successful updates", success, total)
        # Perf summary
        def fmt_s(x):
            return f"{x:.3f}s"
        def fmt_avg(total, count):
            return fmt_s(total / count) if count else "-"
        logger.info(
            "Perf summary | total=%s, login=%s (%.0f%%), search: n=%d total=%s avg=%s min=%s max=%s fail=%d | "
            "save: n=%d total=%s avg=%s min=%s max=%s fail=%d timeouts=%d retries=%d timeout_retries=%d | skipped=%d, vendor_fallback=%d, workers=%d",
            fmt_s(total_elapsed), fmt_s(stats['login_s']), (stats['login_s'] / total_elapsed * 100) if total_elapsed else 0,
            stats['search_count'], fmt_s(stats['search_s_total']), fmt_avg(stats['search_s_total'], stats['search_count']),
            fmt_s(stats['search_s_min'] if stats['search_s_min'] != float('inf') else 0.0), fmt_s(stats['search_s_max']), stats['search_fail'],
            stats['save_count'], fmt_s(stats['save_s_total']), fmt_avg(stats['save_s_total'], stats['save_count']),
            fmt_s(stats['save_s_min'] if stats['save_s_min'] != float('inf') else 0.0), fmt_s(stats['save_s_max']), stats['save_fail'], stats['save_timeouts'],
            stats['save_conflict_retries'], stats['save_timeout_retries'], stats['skipped_unchanged'], stats['vendor_fallback'], max_workers
        )
    return success, total

def load_conf(conf_file):
    with open(conf_file, 'r') as cf:
        conf = yaml.load(cf, Loader=yaml.FullLoader)
        return conf

if __name__ == '__main__':
    # Basic console logging configuration for CLI runs
    logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')

    conf_name = Path('updatecustomfields.yaml')

    if conf_name.is_file():
        conf = load_conf(conf_name)
    else:
        logger.error('Conf file %s not exists!', conf_name)
        sys.exit(1)

    # Izbor Excel fajla
    excel_path = choose_excel_file()
    if not excel_path:
        logger.error("Nije izabran Excel fajl.")
        sys.exit(1)

    try:
        df = pd.read_excel(
            excel_path,
            usecols=["tp_ID", "Godina"],
            dtype={"tp_ID": str, "Godina": str},
            engine="openpyxl"
        )
    except Exception:
        logger.error("Excel fajl mora da sadrži kolone 'tp_ID' i 'Godina'.")
        sys.exit(1)

    # Očisti prazne redove i beline
    df = df.dropna(subset=["tp_ID", "Godina"]).copy()
    df["tp_ID"] = df["tp_ID"].astype(str).str.strip()
    df["Godina"] = df["Godina"].astype(str).str.strip()

    # Priprema podataka za ažuriranje
    live_requests = dict(zip(df['tp_ID'], df['Godina']))

    # Unos korisničkog imena i lozinke
    username = input("Unesite korisničko ime: ")
    pwd = input("Unesite lozinku: ")

    update_live_requests(conf, username, pwd, live_requests)

