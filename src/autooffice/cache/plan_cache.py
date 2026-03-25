"""Plan 캐싱 시스템.

동일 양식 + 동일 작업 유형의 반복 작업에서
Claude 호출 없이 기존 plan을 재사용할 수 있게 한다.

캐시 키: (template_hash, task_type) 조합
저장 위치: ~/.autooffice/cache/
"""

from __future__ import annotations

import hashlib
import json
import logging
from pathlib import Path
from typing import Any

from autooffice.models.execution_plan import ExecutionPlan

logger = logging.getLogger(__name__)

DEFAULT_CACHE_DIR = Path.home() / ".autooffice" / "cache"


class PlanCache:
    """execution_plan.json 캐시 관리자.

    같은 양식(template_hash)과 작업 유형(task_type)에 대해
    이전에 생성된 plan을 재사용할 수 있도록 캐싱한다.
    """

    def __init__(self, cache_dir: Path | None = None) -> None:
        self.cache_dir = cache_dir or DEFAULT_CACHE_DIR
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.index_path = self.cache_dir / "index.json"

    def save_plan(self, plan: ExecutionPlan, plan_json: dict) -> str:
        """plan을 캐시에 저장하고 plan_id를 반환한다."""
        plan_id = plan.task_id

        # plan 파일 저장
        plan_file = self.cache_dir / f"{plan_id}.json"
        plan_file.write_text(
            json.dumps(plan_json, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        # 인덱스 업데이트
        index = self._load_index()
        index[plan_id] = {
            "task_id": plan_id,
            "task_type": plan.metadata.task_type,
            "template_hash": plan.metadata.template_hash,
            "created_at": plan.created_at.isoformat(),
            "description": plan.description,
            "file": str(plan_file),
        }
        self._save_index(index)

        logger.info("Plan 캐시 저장: %s", plan_id)
        return plan_id

    def load_plan(self, plan_id: str) -> ExecutionPlan | None:
        """캐시에서 plan을 로드한다."""
        plan_file = self.cache_dir / f"{plan_id}.json"
        if not plan_file.exists():
            return None

        try:
            data = json.loads(plan_file.read_text(encoding="utf-8"))
            return ExecutionPlan.model_validate(data)
        except Exception as e:
            logger.warning("캐시 로드 실패 (%s): %s", plan_id, e)
            return None

    def find_plan(self, template_hash: str, task_type: str) -> ExecutionPlan | None:
        """template_hash + task_type 조합으로 캐시된 plan을 찾는다.

        같은 양식과 같은 작업 유형이면 기존 plan을 재사용할 수 있다.
        """
        index = self._load_index()
        for plan_id, entry in index.items():
            if (
                entry.get("template_hash") == template_hash
                and entry.get("task_type") == task_type
            ):
                logger.info("캐시 히트: %s (hash=%s, type=%s)", plan_id, template_hash[:8], task_type)
                return self.load_plan(plan_id)

        logger.info("캐시 미스: hash=%s, type=%s", template_hash[:8], task_type)
        return None

    def list_plans(self) -> list[dict[str, Any]]:
        """캐시된 plan 목록을 반환한다."""
        index = self._load_index()
        return list(index.values())

    def invalidate(self, plan_id: str) -> bool:
        """특정 plan을 캐시에서 제거한다."""
        plan_file = self.cache_dir / f"{plan_id}.json"
        if plan_file.exists():
            plan_file.unlink()

        index = self._load_index()
        if plan_id in index:
            del index[plan_id]
            self._save_index(index)
            logger.info("캐시 무효화: %s", plan_id)
            return True
        return False

    def _load_index(self) -> dict[str, Any]:
        if self.index_path.exists():
            return json.loads(self.index_path.read_text(encoding="utf-8"))
        return {}

    def _save_index(self, index: dict) -> None:
        self.index_path.write_text(
            json.dumps(index, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    @staticmethod
    def compute_template_hash(file_path: str | Path) -> str:
        """양식 파일의 SHA256 해시를 계산한다."""
        path = Path(file_path)
        sha256 = hashlib.sha256()
        sha256.update(path.read_bytes())
        return sha256.hexdigest()
