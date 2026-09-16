# -*- coding: utf-8 -*-
from .academic import TEMPLATE as ACADEMIC
from .formal import TEMPLATE as FORMAL
from .clean import TEMPLATE as CLEAN

TEMPLATES = {"academic": ACADEMIC, "formal": FORMAL, "clean": CLEAN}


def get_template(name: str) -> dict:
    tpl = TEMPLATES.get(name)
    if tpl is None:
        raise KeyError(f"unknown template: {name}")
    return tpl
