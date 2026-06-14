template_path = ""
doc = None

def set_template(path, document):
    global template_path, doc
    template_path = path
    doc = document

def get_template():
    global template_path, doc
    return template_path, doc