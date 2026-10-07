"""M23 的注册与编排回归；实际单守卫判别由缺陷用例和变异进程验证。"""
from __future__ import annotations

import pytest

from scripts import mutate_python_docker_defects2 as mutations


@pytest.mark.parametrize("which, other", [("M23a", "M23b"), ("M23b", "M23a")])
def test_input_mutation_removes_only_its_own_guard(which, other):
    source = mutations.RUNNER.read_text(encoding="utf-8")
    assert source.count(mutations.SEEDS[which]) == 1
    mutated = source.replace(mutations.SEEDS[which], mutations.MUTANTS[which])
    assert mutations.SEEDS[which] not in mutated
    assert mutated.count(mutations.SEEDS[other]) == 1
    assert mutations.TARGETS[which][1] != mutations.TARGETS[other][1]


def test_m23_alias_runs_independent_in_memory_mutations_without_source_writes(monkeypatch):
    original = mutations.RUNNER.read_bytes()
    calls = []
    monkeypatch.setattr(mutations, "_install_signal_hooks", lambda: None)

    def fake_pytest(test_path, name, *, runner_source=None):
        calls.append((name, runner_source))
        if runner_source is not None:
            which = name.removesuffix("_mutated")
            other = "M23b" if which == "M23a" else "M23a"
            assert mutations.SEEDS[which] not in runner_source
            assert runner_source.count(mutations.SEEDS[other]) == 1
            test_name = mutations.TARGETS[which][1][0]
            return f"FAILED {test_path}::{test_name}\n1 failed, 1 passed"
        return "2 passed"

    monkeypatch.setattr(mutations, "run_pytest", fake_pytest)
    assert mutations.main("M23") == 0
    assert [name for name, source in calls if source is not None] == [
        "M23a_mutated", "M23b_mutated"
    ]
    assert mutations.RUNNER.read_bytes() == original
    assert mutations._touched == {}


def test_m23_alias_stops_when_first_guard_has_no_discriminating_failure(monkeypatch):
    calls = []
    monkeypatch.setattr(mutations, "_install_signal_hooks", lambda: None)

    def fake_pytest(test_path, name, *, runner_source=None):
        calls.append(name)
        return "2 passed"

    monkeypatch.setattr(mutations, "run_pytest", fake_pytest)
    assert mutations.main("M23") == 5
    assert "M23a_mutated" in calls
    assert "M23b_mutated" not in calls
