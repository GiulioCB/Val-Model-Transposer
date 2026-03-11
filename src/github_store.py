from __future__ import annotations

import base64
import os
from typing import Tuple

from github import Github, Auth
from github.GithubException import GithubException, UnknownObjectException


def _get_repo(repo_name: str, require_token: bool = True):
    token = os.getenv("GITHUB_TOKEN")
    if require_token and not token:
        raise ValueError("GITHUB_TOKEN is not set. Add it as an environment variable before using GitHub storage.")

    client = Github(auth=Auth.Token(token)) if token else Github()
    return client.get_repo(repo_name)


def download_file_bytes(repo_name: str, path: str, branch: str | None = None) -> Tuple[bytes, str]:
    repo = _get_repo(repo_name, require_token=False)
    file = repo.get_contents(path, ref=branch) if branch else repo.get_contents(path)
    return base64.b64decode(file.content), file.sha


def upload_file_bytes(
    repo_name: str,
    path: str,
    content: bytes,
    sha: str | None,
    message: str,
    branch: str | None = None,
) -> None:
    repo = _get_repo(repo_name, require_token=True)
    encoded = base64.b64encode(content).decode()

    if sha:
        repo.update_file(path=path, message=message, content=encoded, sha=sha, branch=branch)
        return

    try:
        existing = repo.get_contents(path, ref=branch) if branch else repo.get_contents(path)
        repo.update_file(path=path, message=message, content=encoded, sha=existing.sha, branch=branch)
    except UnknownObjectException:
        repo.create_file(path=path, message=message, content=encoded, branch=branch)
    except GithubException:
        raise
