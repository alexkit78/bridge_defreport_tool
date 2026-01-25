#project_model.py

def make_empty_project():
    return {
        "bridge": {}, #Form 1
        "spans": [], #Form 2
        "piers": [], #Form 3
        "defects": [], #Form 5
        "photos": { #Photos
            "folder": "",
            "cover": {
                "filename": "",
                "caption" : ""
            },
            "gallery": []
        }
    } #Form 5