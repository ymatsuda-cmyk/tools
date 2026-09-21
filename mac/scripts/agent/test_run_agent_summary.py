import json
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import run_agent


class ReadClineSummaryTest(unittest.TestCase):
    def test_prefers_cline_summary_text(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            repository = Path(temporary_directory)
            (repository / ".agent-summary.json").write_text(
                json.dumps(
                    {
                        "issueNumber": 59,
                        "summaryText": "Issue #59 の実装が完了しました。\n確認結果: テスト済み",
                        "implementation": ["旧形式の内容"],
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )

            with mock.patch.object(run_agent, "TARGET_REPO", repository):
                summary = run_agent.read_cline_summary(59)

        self.assertEqual(
            summary,
            "Issue #59 の実装が完了しました。\n確認結果: テスト済み",
        )

    def test_formats_legacy_section_fields(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            repository = Path(temporary_directory)
            (repository / ".agent-summary.json").write_text(
                json.dumps(
                    {
                        "issueNumber": 59,
                        "implementation": ["カスタマイズタブを追加"],
                        "verification": ["表示を確認"],
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )

            with mock.patch.object(run_agent, "TARGET_REPO", repository):
                summary = run_agent.read_cline_summary(59)

        self.assertEqual(
            summary,
            "【実装内容】\n・カスタマイズタブを追加\n\n"
            "【確認結果】\n・表示を確認",
        )


if __name__ == "__main__":
    unittest.main()