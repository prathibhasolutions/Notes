import json
import os
from typing import Any, Dict, List


class JsonFileStorage:
    """File handling as a lightweight database layer."""

    def __init__(self, data_dir: str):
        self.data_dir = data_dir
        os.makedirs(self.data_dir, exist_ok=True)

    def _path(self, filename: str) -> str:
        return os.path.join(self.data_dir, filename)

    def save_json(self, filename: str, payload: Dict[str, Any]):
        with open(self._path(filename), 'w', encoding='utf-8') as file:
            json.dump(payload, file, indent=4)

    def load_json(self, filename: str):
        path = self._path(filename)
        if not os.path.exists(path):
            return {}
        with open(path, 'r', encoding='utf-8') as file:
            return json.load(file)

    def save_school(self, school_data: Dict[str, Any]):
        self.save_json('school.json', school_data)

    def load_school(self):
        return self.load_json('school.json')

    def list_files(self):
        return [name for name in os.listdir(self.data_dir) if os.path.isfile(os.path.join(self.data_dir, name))]
