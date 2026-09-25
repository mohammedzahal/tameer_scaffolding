from flask import Flask, render_template, request, redirect, url_for, session, send_file, flash, jsonify
import sqlite3
import json
import os
import io
from datetime import datetime

from database import (
    get_db, init_db, get_next_quotation_ref, DEFAULT_TERMS, DEFAULT_BANK_DETAILS,
    DEFAULT_SIGNATORY, DEFAULT_SIGNATORY_TITLE
)
from generate_word import build_quote_word
from generate_excel import build_quote_excel
from generate_pdf import build_quote_pdf

app = Flask(__name__)
app.secret_key = 'tameer_scaffolding_enterprise_secret_key'

# Ensure database is initialized on startup
init_db()

def get_project_dict(row):
    """Converts a database row into a standardized dictionary."""
    p = dict(row)
    try:
        items = json.loads(p.get('items_json') or '[]')
    except:
        items = []
    p['items'] = items
    return p

# --- DASHBOARD ---
@app.route('/')
def home():
    conn = get_db()
    cursor = conn.cursor()

    cursor.execute('SELECT * FROM Projects ORDER BY project_id DESC')
    all_projects = [get_project_dict(r) for r in cursor.fetchall()]

    leads = [p for p in all_projects if p['project_status'] == 'Lead']
    quotes_sent = [p for p in all_projects if p['project_status'] == 'Quote Sent']

    total_pipeline_val = sum([p['grand_total'] for p in quotes_sent if p['grand_total']])

    cursor.execute('SELECT COUNT(*) FROM PriceList')
    catalog_count = cursor.fetchone()[0]

    conn.close()

    return render_template(
        'index.html',
        leads=leads,
        quotes_sent=quotes_sent,
        active_leads_count=len(leads),
        quotes_sent_count=len(quotes_sent),
        total_pipeline_val=total_pipeline_val,
        catalog_count=catalog_count
    )

# --- LEAD MANAGEMENT ---
@app.route('/add_lead', methods=['POST'])
def add_lead():
    conn = get_db()
    cursor = conn.cursor()

    company_name = request.form.get('company_name', '').strip()
    site_name = request.form.get('site_name', '').strip()
    location = request.form.get('location', 'Kingdom of Saudi Arabia').strip()
    attn = request.form.get('attn', '').strip()
    tel = request.form.get('tel', '').strip()
    email = request.form.get('email', '').strip()

    # Generate next official quotation ref: TMR-FZ-XXXX-55471-YYYY
    default_ref = get_next_quotation_ref()
    today_str = datetime.now().strftime("%d %b %Y")

    cursor.execute('''
    INSERT INTO Projects (company_name, site_name, location, attn, tel, email, project_status, quotation_ref, quotation_date)
    VALUES (?, ?, ?, ?, ?, ?, 'Lead', ?, ?)
    ''', (company_name, site_name, location, attn, tel, email, default_ref, today_str))

    project_id = cursor.lastrowid
    conn.commit()
    conn.close()

    flash(f"Lead for '{company_name}' created with Ref: {default_ref}.", "success")
    return redirect(url_for('quote_builder', project_id=project_id))

# --- QUOTE BUILDER & EDITOR ---
@app.route('/quote/<int:project_id>')
def quote_builder(project_id):
    conn = get_db()
    cursor = conn.cursor()

    cursor.execute('SELECT * FROM Projects WHERE project_id = ?', (project_id,))
    row = cursor.fetchone()
    conn.close()

    if not row:
        flash("Quotation project not found.", "danger")
        return redirect(url_for('home'))

    project = get_project_dict(row)
    current_date = datetime.now().strftime("%d %b %Y")
    default_ref = project.get('quotation_ref') or get_next_quotation_ref()

    return render_template(
        'quote_builder.html',
        project=project,
        current_date=current_date,
        default_ref=default_ref,
        default_terms=DEFAULT_TERMS,
        default_bank_details=DEFAULT_BANK_DETAILS,
        default_signatory=DEFAULT_SIGNATORY,
        default_signatory_title=DEFAULT_SIGNATORY_TITLE
    )

@app.route('/save_quote/<int:project_id>', methods=['POST'])
def save_quote(project_id):
    conn = get_db()
    cursor = conn.cursor()

    company_name = request.form.get('company_name', '').strip()
    site_name = request.form.get('site_name', '').strip()
    location = request.form.get('location', '').strip()
    attn = request.form.get('attn', '').strip()
    yr_ref = request.form.get('yr_ref', '').strip()
    quotation_ref = request.form.get('quotation_ref', '').strip()
    quotation_date = request.form.get('quotation_date', '').strip()
    subject = request.form.get('subject', '').strip()
    intro_note = request.form.get('intro_note', '').strip()
    terms_conditions = request.form.get('terms_conditions', '').strip()
    bank_details = request.form.get('bank_details', '').strip()
    signatory_name = request.form.get('signatory_name', '').strip()
    signatory_title = request.form.get('signatory_title', '').strip()

    subtotal = float(request.form.get('subtotal', 0.0) or 0.0)
    vat_amount = float(request.form.get('vat_amount', 0.0) or 0.0)
    grand_total = float(request.form.get('grand_total', 0.0) or 0.0)
    items_json = request.form.get('items_json', '[]')

    cursor.execute('''
    UPDATE Projects SET 
        company_name = ?,
        site_name = ?,
        location = ?,
        attn = ?,
        quotation_ref = ?,
        quotation_date = ?,
        subject = ?,
        intro_note = ?,
        terms_conditions = ?,
        bank_details = ?,
        signatory_name = ?,
        signatory_title = ?,
        subtotal = ?,
        vat_amount = ?,
        grand_total = ?,
        items_json = ?,
        project_status = 'Quote Sent'
    WHERE project_id = ?
    ''', (
        company_name, site_name, location, attn, quotation_ref,
        quotation_date, subject, intro_note, terms_conditions,
        bank_details, signatory_name, signatory_title, subtotal,
        vat_amount, grand_total, items_json, project_id
    ))

    conn.commit()
    conn.close()

    flash(f"Quotation '{quotation_ref}' updated & calculated successfully! Total: SAR {grand_total:,.2f}", "success")
    return redirect(url_for('quote_builder', project_id=project_id))

# --- EXPORT & DOWNLOAD ROUTES ---
@app.route('/download_word/<int:project_id>')
def download_word(project_id):
    try:
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM Projects WHERE project_id = ?', (project_id,))
        row = cursor.fetchone()
        conn.close()

        if not row:
            flash("Quotation not found.", "danger")
            return redirect(url_for('home'))

        quote_data = get_project_dict(row)
        docx_buf = build_quote_word(quote_data)
        ref_safe = (quote_data.get('quotation_ref') or f"Q{project_id}").replace('/', '-').replace(' ', '_')
        filename = f"Quotation_{ref_safe}.docx"

        return send_file(
            docx_buf,
            download_name=filename,
            as_attachment=True,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        app.logger.error(f"Error generating Word file for project {project_id}: {e}", exc_info=True)
        flash(f"Could not generate Word document: {str(e)}", "danger")
        return redirect(url_for('home'))

@app.route('/download_excel/<int:project_id>')
def download_excel(project_id):
    try:
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM Projects WHERE project_id = ?', (project_id,))
        row = cursor.fetchone()
        conn.close()

        if not row:
            flash("Quotation not found.", "danger")
            return redirect(url_for('home'))

        quote_data = get_project_dict(row)
        xlsx_buf = build_quote_excel(quote_data)
        ref_safe = (quote_data.get('quotation_ref') or f"Q{project_id}").replace('/', '-').replace(' ', '_')
        filename = f"Quotation_{ref_safe}.xlsx"

        return send_file(
            xlsx_buf,
            download_name=filename,
            as_attachment=True,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        app.logger.error(f"Error generating Excel file for project {project_id}: {e}", exc_info=True)
        flash(f"Could not generate Excel spreadsheet: {str(e)}", "danger")
        return redirect(url_for('home'))

@app.route('/download_pdf/<int:project_id>')
def download_pdf(project_id):
    try:
        conn = get_db()
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM Projects WHERE project_id = ?', (project_id,))
        row = cursor.fetchone()
        conn.close()

        if not row:
            flash("Quotation not found.", "danger")
            return redirect(url_for('home'))

        quote_data = get_project_dict(row)
        pdf_buf = build_quote_pdf(quote_data)
        ref_safe = (quote_data.get('quotation_ref') or f"Q{project_id}").replace('/', '-').replace(' ', '_')
        filename = f"Quotation_{ref_safe}.pdf"

        return send_file(
            pdf_buf,
            download_name=filename,
            as_attachment=True,
            mimetype='application/pdf'
        )
    except Exception as e:
        app.logger.error(f"Error generating PDF file for project {project_id}: {e}", exc_info=True)
        flash(f"Could not generate PDF document: {str(e)}", "danger")
        return redirect(url_for('home'))

@app.route('/print_quote/<int:project_id>')
def print_quote(project_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT * FROM Projects WHERE project_id = ?', (project_id,))
    row = cursor.fetchone()
    conn.close()

    if not row:
        flash("Quotation not found", "danger")
        return redirect(url_for('home'))

    quote_data = get_project_dict(row)
    return render_template('quote_print.html', project=quote_data, items=quote_data.get('items', []))

@app.route('/delete_project/<int:project_id>', methods=['POST'])
def delete_project(project_id):
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM Projects WHERE project_id = ?', (project_id,))
    conn.commit()
    conn.close()

    flash("Project deleted successfully.", "info")
    return redirect(url_for('home'))

# --- CATALOG & PRICE SETTINGS ---
@app.route('/settings')
def settings():
    conn = get_db()
    cursor = conn.cursor()

    cursor.execute('SELECT * FROM PriceList ORDER BY category, item_number, item_name')
    items = [dict(r) for r in cursor.fetchall()]

    cursor.execute('SELECT DISTINCT category FROM PriceList ORDER BY category')
    categories = [r[0] for r in cursor.fetchall()]

    conn.close()
    return render_template('settings.html', items=items, categories=categories)

@app.route('/api/items')
def api_items():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT item_id, item_number, category, item_name, unit, unit_price FROM PriceList ORDER BY item_name')
    items = [dict(r) for r in cursor.fetchall()]
    conn.close()
    return jsonify(items)

@app.route('/api/update_price', methods=['POST'])
def api_update_price():
    data = request.get_json() or {}
    item_id = data.get('item_id')
    unit_price = data.get('unit_price')

    if item_id is None or unit_price is None:
        return jsonify({'error': 'Invalid parameters'}), 400

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('UPDATE PriceList SET unit_price = ?, updated_at = CURRENT_TIMESTAMP WHERE item_id = ?',
                   (float(unit_price), int(item_id)))
    conn.commit()
    conn.close()

    return jsonify({'success': True, 'item_id': item_id, 'unit_price': float(unit_price)})

@app.route('/api/add_item', methods=['POST'])
def api_add_item():
    category = request.form.get('category', 'General').strip()
    item_name = request.form.get('item_name', '').strip()
    unit = request.form.get('unit', 'Pcs.').strip()
    unit_price = float(request.form.get('unit_price', 0.0) or 0.0)

    if not item_name:
        flash("Item name is required.", "danger")
        return redirect(url_for('settings'))

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('SELECT MAX(item_number) FROM PriceList')
    max_num = cursor.fetchone()[0] or 136
    new_num = max_num + 1

    cursor.execute('''
    INSERT INTO PriceList (item_number, category, item_name, unit, unit_price)
    VALUES (?, ?, ?, ?, ?)
    ''', (new_num, category, item_name, unit, unit_price))
    conn.commit()
    conn.close()

    flash(f"Added component '{item_name}' (SAR {unit_price:.2f}) to catalog.", "success")
    return redirect(url_for('settings'))

if __name__ == '__main__':
    # Running on port 5001 to keep port 5000 free
    app.run(host='0.0.0.0', port=5001, debug=True)
