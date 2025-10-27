import os
import re
import shutil
import json
import tempfile
from concurrent.futures import ThreadPoolExecutor, as_completed

import vpsdb
import git
from github import Github, Auth
from github.GithubException import GithubException
from pathlib import Path

def find_table_yml(base_dir="external"):
    result = []
    if not os.path.exists(base_dir):
        print(f"Directory {base_dir} does not exist.")
        return result
    for entry in os.listdir(base_dir):
        entry_path = os.path.join(base_dir, entry)
        if os.path.isdir(entry_path) and entry.startswith("vpx-"):
            table_yml = os.path.join(entry_path, "table.yml")
            if os.path.exists(table_yml):
                result.append(table_yml)
    return result

def get_latest_commit_hash(repo_path, folder_path):
    """
    Retrieves the latest commit hash that modified files in a specific folder.
    
	Args:
        repo_path (str): The path to the local Git repository.
        folder_path (str): The folder path within the repository.

    Returns:
        str: The latest commit hash, or None if no matching commits are found.
    """
    try:
        repo = git.Repo(repo_path)
        commits = list(repo.iter_commits(paths=folder_path, max_count=1))  # Only the most recent commit
        if commits:
            return commits[0].hexsha
        else:
            return None
    except git.InvalidGitRepositoryError:
        print(f"Error: Invalid Git repository at {repo_path}")
        return None
    except Exception as e:
        print(f"An error occurred: {e}")
        return None

def process_title(title, manufacturer, year):
    """
    Transforms a title for proper sorting, moving leading "The" to the end,
    and handling optional "JP's" or "JPs" prefixes, moving them after the comma
    when 'The' is not present.
    """
    name = ""
    match_the = re.match(r"^(JP'?s\s*)?(The)\s+(.+)$", title)
    match_jps = re.match(r"^(JP'?s)\s+(.+)$", title)
    if match_the and match_the.group(2):
        prefix = match_the.group(1) or ""
        name = f"{match_the.group(3)}, {prefix}{match_the.group(2)}"
    elif match_jps:
        name = f"{match_jps.group(2)}, {match_jps.group(1)}"
    else:
        name = title
    return f"{name} ({manufacturer} {year})"

import time
from github.GithubException import GithubException

def upload_release_asset(github_token, repo_name, release_tag, file_path, clobber=True, max_attempts=3):
    """
    Uploads a file as a release asset using PyGithub. Returns the browser_download_url or None.

    Common 403 causes handled here:
    - "Resource not accessible by integration" => GITHUB_TOKEN lacks contents:write or repo setting is read-only
    - Wrong repo (GITHUB_TOKEN can't write across repos)
    - Race while deleting/creating assets
    """
    file_name = os.path.basename(file_path)
    try:
        auth = Auth.Token(github_token)
        g = Github(auth=auth)
        repo = g.get_repo(repo_name)

        # Sanity: fetch the release by tag, raise clear error if missing
        try:
            release = repo.get_release(release_tag)
        except GithubException as ge:
            if ge.status == 404:
                print(f"[ERROR] Release with tag '{release_tag}' not found in {repo_name}.")
            else:
                print(f"[ERROR] Could not read release '{release_tag}' ({ge.status}): {ge.data}")
            return None

        # Optional clobber of existing asset (by name)
        if clobber:
            try:
                for asset in release.get_assets():
                    if asset.name == file_name:
                        print(f"[INFO] Deleting existing asset '{file_name}'...")
                        asset.delete_asset()
                        # tiny delay to avoid race with eventual consistency
                        time.sleep(0.8)
                        break
            except GithubException as ge:
                print(f"[WARN] Could not list/delete existing assets ({ge.status}): {ge.data}")

        # Upload with correct content type (zip) and stable name
        attempt = 0
        while attempt < max_attempts:
            attempt += 1
            try:
                print(f"[INFO] Uploading '{file_name}' (attempt {attempt}/{max_attempts})...")
                # PyGithub signature: upload_asset(path, label=None, name=None, content_type='application/octet-stream')
                asset = release.upload_asset(
                    file_path,
                    label=file_name,
                    name=file_name,
                    content_type="application/zip" if file_name.endswith(".zip") else "application/octet-stream",
                )
                print(f"[INFO] Uploaded {file_name} to release {release_tag}")

                # Refresh and return URL
                release = repo.get_release(release_tag)
                for a in release.get_assets():
                    if a.name == file_name:
                        return a.browser_download_url
                print(f"[WARN] Uploaded asset '{file_name}' not visible yet; retrying index refresh...")
                time.sleep(1.0)
            except GithubException as ge:
                # 403s need actionable messages
                if ge.status == 403:
                    msg = ge.data if isinstance(ge.data, dict) else str(ge.data)
                    print(f"[ERROR] 403 Forbidden while uploading '{file_name}': {msg}")
                    print(
                        "HINTS: "
                        "1) Ensure workflow has `permissions: contents: write` "
                        "2) Ensure repo setting 'Workflow permissions' is 'Read and write' "
                        "3) Ensure you're uploading to the SAME repo the workflow runs in "
                        "4) Fine-grained PAT needed if targeting a different repo"
                    )
                    # 403 is usually not transient—break
                    break
                elif ge.status in (502, 503, 504):
                    print(f"[WARN] Transient server error ({ge.status}); will retry...")
                    time.sleep(2 ** attempt)
                else:
                    print(f"[ERROR] Upload failed ({ge.status}): {ge.data}")
                    break
            except Exception as e:
                print(f"[ERROR] Unexpected upload error: {e}")
                break

        return None
    except Exception as e:
        print(f"[ERROR] upload_release_asset fatal: {e}")
        return None


def process_table(args):
    # Unpack for ThreadPoolExecutor compatibility
    table, table_data, github_token, repo_name, release_tag = args
    external_path = os.path.join("external", table)
    result = table, table_data.copy()
    if not os.path.isdir(external_path):
        print(f"Warning: Directory {external_path} does not exist, skipping {table}.")
        return result

    # Get latest commit for this folder
    config_version = get_latest_commit_hash(".", external_path)
    if not config_version:
        print(f"Error: No commit found for {external_path}, skipping {table}.")
        return result

    # Optional: Skip if unchanged (requires storing previous version info in manifest)
    # Uncomment below if you want to skip tables whose hash matches what's already in manifest:
    # if table_data.get("configVersion") == config_version[:7]:
    #     print(f"Skipping unchanged table: {table}")
    #     return result

    # Zip the table directory in a temp folder
    with tempfile.TemporaryDirectory() as tmpdir:
        zip_base = os.path.join(tmpdir, table)
        try:
            shutil.make_archive(zip_base, "zip", external_path)
            zip_path = zip_base + ".zip"
            print(f"Uploading {zip_path} to GitHub...")
            download_url = upload_release_asset(
                github_token, repo_name, release_tag, zip_path
            )
            if download_url:
                new_data = result[1]
                new_data["repoConfig"] = download_url
                new_data["configVersion"] = config_version[:7]
                new_data["name"] = process_title(new_data["name"], new_data["manufacturer"], new_data["year"])
                result = (table, new_data)
            else:
                print(f"Failed to upload asset for {table}")
        except Exception as e:
            print(f"Error processing {table}: {e}")
    return result


if __name__ == "__main__":
    github_token = os.environ.get("GITHUB_TOKEN")
    repo_name = os.environ.get("GITHUB_REPOSITORY")
    release_tag = os.environ.get("GITHUB_REF_NAME")

    if not github_token or not repo_name or not release_tag:
        print("Error: Required environment variables not set.")
        sys.exit(1)

    try:
        g = Github(auth=Auth.Token(github_token))
        repo = g.get_repo(repo_name)
        rel = repo.get_release(release_tag)

        # Capability probe: listing assets should succeed with a write-capable token
        _ = list(rel.get_assets())

    except GithubException as ge:
        if ge.status == 403:
            print("[ERROR] Token cannot access the release. Likely missing 'contents: write' or repo workflow perms set to read-only.")
        elif ge.status == 404:
            print(f"[ERROR] Release '{release_tag}' not found in '{repo_name}'.")
        else:
            print(f"[ERROR] Unable to access release ({ge.status}): {ge.data}")
        sys.exit(1)

    except Exception as e:
        print(f"[ERROR] Unexpected error while probing release access: {e}")
        sys.exit(1)
	
    files = find_table_yml()
    tables = vpsdb.get_table_meta(files)

    # Remove disabled tables before processing
    tables_to_remove = [table for table, data in tables.items() if data.get("enabled") is False]
    for table in tables_to_remove:
        print(f"Skipping disabled table: {table}")
        del tables[table]

    # Prepare arguments for parallel processing
    pool_args = [
        (table, tables[table], github_token, repo_name, release_tag)
        for table in tables
    ]

    # Process tables in parallel (adjust max_workers as needed)
    updated_tables = {}
    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = {executor.submit(process_table, arg): arg[0] for arg in pool_args}
        for future in as_completed(futures):
            table, updated_data = future.result()
            updated_tables[table] = updated_data

    # Write updated manifest
    manifest_file = "manifest.json"
    with open(manifest_file, "w") as f:
        json.dump(updated_tables, f, indent=2)

    # Upload manifest.json to release
    manifest_url = upload_release_asset(github_token, repo_name, release_tag, manifest_file)
    print(f"Uploaded manifest.json to release: {manifest_url}")

    # Optional: remove manifest after upload
    os.remove(manifest_file)
