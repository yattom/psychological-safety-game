#coding: utf8

import pytest

from cardgen import main

PRESENTATION_ID = '125AtwimPYNK4gGO0aDwQ1NXxU4kuoR84D7cvuLpaRy8'


def test_detect_elements_to_modify():
    credentials = main.auth()
    service = main.create_service_for_slides(credentials)
    presentation = service.presentations().get(presentationId=PRESENTATION_ID).execute()

    elements_to_modify = main.detect_elements_to_modify(['テスト項目1'], presentation)
    assert 'テスト項目1' in elements_to_modify

