"""
Gestor Financeiro de Condomínios — API Web v2.0 (Fase 2: Supabase)
Backend Flask para deploy no Render.

Endpoints:
  GET  /api/health      → ping (aquecimento do free tier) — público
  POST /api/analisar    → recebe planilha Excel, devolve JSON da análise — autenticado
  POST /api/gerar-pptx  → recebe planilha, devolve .pptx e grava histórico — autenticado

Autenticação:
  Header "Authorization: Bearer <access_token do Supabase Auth>".
  O token é validado contra o endpoint /auth/v1/user do Supabase — funciona
  tanto com chaves legadas (JWT HS256) quanto com o formato novo
  (sb_publishable_...), sem guardar segredos de assinatura no backend.

Variáveis de ambiente:
  SUPABASE_URL       ex.: https://xxxx.supabase.co
  SUPABASE_ANON_KEY  chave publishable/anon do MESMO projeto
  CORS_ORIGINS       ex.: https://seusite.netlify.app
  (Se SUPABASE_URL não estiver definida, a API roda em MODO ABERTO —
   apenas para desenvolvimento local. Em produção, configure sempre.)
"""

import os
import re
import tempfile
import unicodedata
from functools import wraps
from io import BytesIO
from datetime import datetime

import numpy as np
import requests as http
from flask import Flask, request, jsonify, send_file, g
from flask_cors import CORS

from condominio_app_v2 import CondominiumFinancialAnalyzer
from powerpoint_generator import PowerPointGenerator

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 10 * 1024 * 1024  # 10 MB

_origins = os.environ.get('CORS_ORIGINS', '*')
CORS(app, origins=_origins.split(',') if _origins != '*' else '*')

SUPABASE_URL      = (os.environ.get('https://scjiztniwqmdpzjppigm.supabase.co') or '').rstrip('/')
SUPABASE_ANON_KEY = os.environ.get('eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InNjaml6dG5pd3FtZHB6anBwaWdtIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODMwMTA4NDYsImV4cCI6MjA5ODU4Njg0Nn0.kXIi6iBxyJ7YVI9LjUHGU9Zjct9MtnNwtQMgJJ94cVM') or ''
AUTH_ATIVO        = bool(SUPABASE_URL and SUPABASE_ANON_KEY)

if not AUTH_ATIVO:
    print('⚠  SUPABASE_URL/SUPABASE_ANON_KEY não configuradas — '
          'API rodando em MODO ABERTO (apenas desenvolvimento).')

EXTENSOES_PERMITIDAS = {'.xls', '.xlsx'}


# ─────────────────────────────────────────────────────────────────────────────
# AUTENTICAÇÃO (Supabase)
# ─────────────────────────────────────────────────────────────────────────────
def requer_auth(f):
    """Valida o access_token do Supabase. Em modo aberto, deixa passar."""
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not AUTH_ATIVO:
            g.user_token = None
            g.user_id = None
            return f(*args, **kwargs)

        auth_header = request.headers.get('Authorization', '')
        if not auth_header.startswith('Bearer '):
            return jsonify({'erro': 'Não autenticado. Faça login para usar o serviço.'}), 401
        token = auth_header[7:].strip()

        try:
            resp = http.get(
                f'{SUPABASE_URL}/auth/v1/user',
                headers={'Authorization': f'Bearer {token}', 'apikey': SUPABASE_ANON_KEY},
                timeout=10,
            )
        except http.RequestException:
            return jsonify({'erro': 'Falha ao validar a sessão. Tente novamente.'}), 503

        if resp.status_code != 200:
            return jsonify({'erro': 'Sessão inválida ou expirada. Faça login novamente.'}), 401

        g.user_token = token
        g.user_id = resp.json().get('id')
        return f(*args, **kwargs)
    return wrapper


def _gravar_historico(r: dict, nome_arquivo: str):
    """
    Insere o registro no histórico via PostgREST usando o token DO USUÁRIO,
    respeitando o RLS (user_id preenchido por default auth.uid() no banco).
    Falha aqui não impede a entrega do PPTX — só registra no log.
    """
    if not (AUTH_ATIVO and g.get('user_token')):
        return
    payload = {
        'nome_condominio':   r.get('nome_condominio'),
        'data_assembleia':   r.get('data_assembleia'),
        'ano_previsao':      _num(r.get('ano_proximo'), int),
        'num_unidades':      _num(r.get('num_unidades'), int),
        'total_despesas':    _num(r.get('total_despesas')),
        'total_rateado':     _num(r.get('total_rateado')),
        'taxa_ideal_mensal': _num(r.get('taxa_ideal_mensal')),
        'taxa_atual':        _num(r.get('taxa_atual')),
        'reajuste_pct':      _num(r.get('reajuste_pct')),
        'nome_arquivo':      nome_arquivo,
    }
    try:
        resp = http.post(
            f'{SUPABASE_URL}/rest/v1/relatorios',
            json=payload,
            headers={
                'Authorization': f'Bearer {g.user_token}',
                'apikey': SUPABASE_ANON_KEY,
                'Content-Type': 'application/json',
                'Prefer': 'return=minimal',
            },
            timeout=10,
        )
        if resp.status_code not in (200, 201):
            print(f'⚠  Histórico não gravado ({resp.status_code}): {resp.text[:200]}')
    except http.RequestException as e:
        print(f'⚠  Histórico não gravado: {e}')


# ─────────────────────────────────────────────────────────────────────────────
# AUXILIARES
# ─────────────────────────────────────────────────────────────────────────────
def _num(v, cast=float):
    """Converte numpy/None para tipo nativo ou None (para o payload do banco)."""
    if v is None:
        return None
    try:
        v = cast(v)
        return None if (isinstance(v, float) and v != v) else v
    except (TypeError, ValueError):
        return None


def _sanitize_json(obj):
    """Converte tipos numpy/pandas para tipos nativos serializáveis em JSON."""
    if isinstance(obj, dict):
        return {k: _sanitize_json(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_sanitize_json(v) for v in obj]
    if isinstance(obj, (np.integer,)):
        return int(obj)
    if isinstance(obj, (np.floating,)):
        v = float(obj)
        return None if np.isnan(v) else v
    if isinstance(obj, np.bool_):
        return bool(obj)
    if isinstance(obj, float) and obj != obj:  # NaN nativo
        return None
    return obj


def _slug_filename(nome: str) -> str:
    """Gera nome de arquivo seguro a partir do nome do condomínio."""
    if not nome:
        return "Relatorio_Condominio"
    s = unicodedata.normalize('NFKD', nome).encode('ascii', 'ignore').decode()
    s = re.sub(r'[^A-Za-z0-9]+', '_', s).strip('_')
    return f"Relatorio_{s}" if s else "Relatorio_Condominio"


def _rodar_analise(req):
    """
    Fluxo compartilhado: valida o upload, roda o analisador e injeta
    os campos opcionais (mesma lógica da GUI desktop v3).
    Retorna (analyzer, None) em sucesso ou (None, (payload, status)) em erro.
    """
    arquivo = req.files.get('planilha')
    if arquivo is None or arquivo.filename == '':
        return None, ({'erro': 'Nenhuma planilha enviada. Envie um arquivo .xls ou .xlsx no campo "planilha".'}, 400)

    ext = os.path.splitext(arquivo.filename)[1].lower()
    if ext not in EXTENSOES_PERMITIDAS:
        return None, ({'erro': f'Formato "{ext}" não suportado. Envie .xls ou .xlsx.'}, 400)

    nome = (req.form.get('nome_condominio') or '').strip()
    data = (req.form.get('data_assembleia') or '').strip()
    taxa = (req.form.get('taxa_atual') or '').strip()

    # O analisador trabalha com caminho de arquivo (detecção de engine por
    # extensão), então gravamos em arquivo temporário com o sufixo correto.
    tmp = tempfile.NamedTemporaryFile(suffix=ext, delete=False)
    try:
        arquivo.save(tmp.name)
        tmp.close()

        analyzer = CondominiumFinancialAnalyzer(tmp.name)
        if analyzer.df_raw is None:
            return None, ({'erro': 'Não foi possível ler a planilha. Verifique se o arquivo está íntegro.'}, 422)

        # Injeta dados opcionais ANTES do processamento (taxa influencia o reajuste)
        if taxa:
            try:
                t = float(taxa.replace('R$', '').replace('.', '').replace(',', '.').strip()) \
                    if ',' in taxa else float(taxa.replace('R$', '').strip())
                analyzer._taxa_atual = t
            except (ValueError, AttributeError):
                pass

        if not analyzer.process_data():
            return None, ({'erro': 'Falha ao processar os dados. Confira se a planilha segue a estrutura de grupos numerados (ex.: "1. Encargos Bancários").'}, 422)

        # Re-injeta após process_data (que pode sobrescrever) — mesma lógica da GUI
        if nome:
            analyzer.analysis_results['nome_condominio'] = nome.upper()
        if data:
            analyzer.analysis_results['data_assembleia'] = data

        return analyzer, None
    finally:
        try:
            os.unlink(tmp.name)
        except OSError:
            pass


# ─────────────────────────────────────────────────────────────────────────────
# ROTAS
# ─────────────────────────────────────────────────────────────────────────────
@app.route('/api/health', methods=['GET'])
def health():
    return jsonify({
        'status': 'ok',
        'servico': 'gestor-condominio-api',
        'versao': '2.0',
        'auth': AUTH_ATIVO,
    })


@app.route('/api/analisar', methods=['POST'])
@requer_auth
def analisar():
    analyzer, erro = _rodar_analise(request)
    if erro:
        payload, status = erro
        return jsonify(payload), status

    r = analyzer.analysis_results
    resumo = _sanitize_json({
        'nome_condominio':   r.get('nome_condominio'),
        'data_assembleia':   r.get('data_assembleia'),
        'ano_proximo':       r.get('ano_proximo'),
        'num_unidades':      r.get('num_unidades'),
        'grupos':            r.get('grupos', []),
        'total_despesas':    r.get('total_despesas'),
        'fundo_reserva':     r.get('fundo_reserva'),
        'fundo_reserva_pct': r.get('fundo_reserva_pct'),
        'garantidora':       r.get('garantidora'),
        'garantidora_pct':   r.get('garantidora_pct'),
        'total_rateado':     r.get('total_rateado'),
        'taxa_ideal_mensal': r.get('taxa_ideal_mensal'),
        'taxa_atual':        r.get('taxa_atual'),
        'reajuste_pct':      r.get('reajuste_pct'),
    })
    return jsonify({'ok': True, 'analise': resumo})


@app.route('/api/gerar-pptx', methods=['POST'])
@requer_auth
def gerar_pptx():
    analyzer, erro = _rodar_analise(request)
    if erro:
        payload, status = erro
        return jsonify(payload), status

    # prs.save() aceita stream: geramos o PPTX inteiramente em memória,
    # sem tocar no filesystem efêmero do Render.
    buffer = BytesIO()
    gen = PowerPointGenerator(analyzer, output_path=buffer)
    gen.generate()
    buffer.seek(0)

    nome_arquivo = _slug_filename(analyzer.analysis_results.get('nome_condominio', ''))
    nome_arquivo += f"_{datetime.now().strftime('%Y%m%d')}.pptx"

    # Grava histórico no Supabase (não bloqueia a entrega em caso de falha)
    _gravar_historico(analyzer.analysis_results, nome_arquivo)

    return send_file(
        buffer,
        as_attachment=True,
        download_name=nome_arquivo,
        mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )


@app.errorhandler(413)
def arquivo_grande(_):
    return jsonify({'erro': 'Arquivo maior que 10 MB. Envie uma planilha menor.'}), 413


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=True)
