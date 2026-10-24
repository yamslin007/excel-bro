"""Result dimensions are needed for planning; result content must stay local."""
import json

import pytest

from server.app.python_analysis import project_python_analysis_receipt


def receipt(card):
    return {'status': 'completed', 'reasonCategory': 'none',
            'resultHandle': 'res_test', 'resultCard': card}


def test_only_registered_result_dimensions_are_projected():
    card = {'resultRef': 'res_test', 'rowCount': 9, 'columnCount': 7,
            'headers': ['PRIVATE_SENTINEL'], 'rows': [['PRIVATE_SENTINEL']],
            'title': 'PRIVATE_SENTINEL', 'sourceSheets': ['PRIVATE_SENTINEL']}
    result = project_python_analysis_receipt(receipt(card))
    assert result['resultShape'] == {'rowCount': 9, 'columnCount': 7}
    assert set(result) == {'status', 'reasonCategory', 'message', 'resultHandle', 'resultShape'}
    assert 'PRIVATE_SENTINEL' not in json.dumps(result)


@pytest.mark.parametrize('field,value', [
    ('rowCount', True), ('rowCount', 0), ('rowCount', '9'),
    ('columnCount', -1), ('columnCount', None), ('resultRef', 'res_other'),
])
def test_invalid_or_unrelated_structure_is_rejected(field, value):
    card = {'resultRef': 'res_test', 'rowCount': 9, 'columnCount': 7, field: value}
    with pytest.raises(ValueError):
        project_python_analysis_receipt(receipt(card))


def test_legacy_receipt_does_not_invent_shape():
    result = project_python_analysis_receipt(receipt(None))
    assert 'resultShape' not in result
