import os
if 'proposal' not in request.files:
flash('No proposal file part')
return redirect(request.url)


proposal_file = request.files['proposal']
notes = request.form.get('notes', '')
template_json = request.form.get('template_json', '')


if proposal_file.filename == '':
flash('No selected file')
return redirect(request.url)


if proposal_file and allowed_file(proposal_file.filename):
filename = secure_filename(proposal_file.filename)
save_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
proposal_file.save(save_path)


# Extract text from PPTX
proposal_text = extract_text_from_pptx(save_path)


# Load template: either user-provided or default
try:
if template_json.strip():
template = json.loads(template_json)
else:
with open('templates_template.json', 'r', encoding='utf-8') as f:
template = json.load(f)
except Exception as e:
flash(f'Template JSON parse error: {e}')
return redirect(url_for('index'))


# Generate sections via LLM
sections = generate_sections_from_template(template, proposal_text, notes)


# Save to docx
out_path = os.path.join(app.config['UPLOAD_FOLDER'], f"Final_Report_{secure_filename(filename)}.docx")
save_sections_to_docx(sections, out_path)


return render_template('result.html', sections=sections, download_path=url_for('download_file', filename=os.path.basename(out_path)))


else:
flash('Allowed file types: pptx')
return redirect(request.url)


@app.route('/download/<filename>')
def download_file(filename):
path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
if not os.path.exists(path):
flash('File not found')
return redirect(url_for('index'))
return send_file(path, as_attachment=True)


if __name__ == '__main__':
app.run(host='0.0.0.0', port=int(os.getenv('PORT', 5000)), debug=True)
