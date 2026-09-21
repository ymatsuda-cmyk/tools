import pathlib
import subprocess
import unittest
from unittest import mock

import run_agent_daemon


class ProcessIssueTest(unittest.TestCase):
    def test_handles_uncaptured_process_output(self) -> None:
        state_path = pathlib.Path("issue-58.json")
        run_agent_path = pathlib.Path("run_agent.py")
        result = subprocess.CompletedProcess(
            args=["python", "run_agent.py", "58"],
            returncode=1,
            stdout=None,
            stderr=None,
        )

        with (
            mock.patch.object(run_agent_daemon.subprocess, "run", return_value=result),
            mock.patch.object(
                run_agent_daemon,
                "load_json_when_ready",
                return_value={"status": "daemon_starting"},
            ),
            mock.patch.object(run_agent_daemon, "update_state") as update_state,
        ):
            exit_code = run_agent_daemon.process_issue(
                58,
                state_path,
                run_agent_path,
                pathlib.Path("repository"),
                no_vscode=True,
                manual=False,
            )

        self.assertEqual(exit_code, 1)
        output_update = update_state.call_args_list[1]
        self.assertEqual(output_update.args[0], state_path)
        self.assertEqual(output_update.args[1]["stdout"], "")
        self.assertEqual(output_update.args[1]["stderr"], "")


if __name__ == "__main__":
    unittest.main()