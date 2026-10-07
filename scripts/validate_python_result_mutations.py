"""Independent process-local mutations; never rewrite dirty production files."""
import argparse
import inspect
import json
import subprocess
import sys


def run_case(case):
    import pytest
    from server.app import python_analysis, python_datasets
    from server.app.python_datasets import PythonDatasetRegistry, DatasetRejected
    if case == "receipt-binding":
        owner,attribute = PythonDatasetRegistry,"_result"
        def mutant(self,token,scope,origin,inputs,receipt):
            grant = self._grant(token,scope,origin)
            result = grant.results.get(receipt.resultHandle)
            if not result or not self.valid(token,scope,origin,inputs): raise DatasetRejected()
            return grant,result
        selector = "results_require_every_binding or exact_order"
    elif case == "matrix-validation":
        owner,attribute = python_datasets,"encode_matrix_result"
        def mutant(result):
            return json.dumps(result,ensure_ascii=False,allow_nan=False,separators=(",",":")).encode("utf-8"),1,1
        selector = "untrusted_output_has_strict_matrix_contract"
    else:
        owner,attribute = python_analysis,"run_registered_result_internal"
        source = inspect.getsource(getattr(owner,attribute))
        guard = "if (cancel_check and cancel_check()) or not data_registry.valid(scope_token, scope, origin, receipts):"
        assert source.count(guard) == 1
        namespace = dict(vars(owner))
        # Developer-owned test mutation only; no model code or cell values.
        exec(compile(source.replace(guard,"if False:"),"<result-cancel-mutant>","exec"),namespace)
        mutant = namespace[attribute]
        selector = "cancel_after_publication"
    original = getattr(owner,attribute)
    try:
        setattr(owner,attribute,mutant)
        result = pytest.main(["server/tests/test_python_results.py","-q","-p","no:cacheprovider","-k",selector])
    finally:
        setattr(owner,attribute,original)
    if result != 1: raise SystemExit("RED-MISS: "+case)
    print("RED-OK: "+case)


def main():
    cases = ("receipt-binding","matrix-validation","late-publication")
    parser = argparse.ArgumentParser();parser.add_argument("--case",choices=cases)
    case = parser.parse_args().case
    if case: return run_case(case)
    for case in cases:
        subprocess.run([sys.executable,"-B","-m","scripts.validate_python_result_mutations","--case",case],check=True)
        subprocess.run([sys.executable,"-B","-m","pytest","server/tests/test_python_results.py","-q","-p","no:cacheprovider"],check=True)
    print("All independent mutants RED; each restored fresh suite GREEN.")


if __name__ == "__main__": main()
