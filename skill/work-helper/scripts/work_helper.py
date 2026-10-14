"""Work-Helper 백엔드를 호출하는 Codex / Claude Code 공용 CLI."""

import argparse
import json
import sys
from pathlib import Path


def parser():
    cli = argparse.ArgumentParser(description=__doc__)
    cli.add_argument("--repo-root", type=Path, default=Path(__file__).resolve().parents[3])
    commands = cli.add_subparsers(dest="command", required=True)

    merge = commands.add_parser("merge", help="xlsx 시트 취합")
    merge.add_argument("--folder", required=True)
    merge.add_argument("--prefix", required=True)
    merge.add_argument("--sheets", required=True)
    merge.add_argument("--headers", default="1")
    merge.add_argument("--recursive", action="store_true")

    for name in ("roster", "recipients", "folder-info", "folders"):
        command = commands.add_parser(name)
        command.add_argument("--file", required=True)
        if name == "recipients":
            command.add_argument("--teams", required=True)
            command.add_argument("--to-roles", default="")
            command.add_argument("--cc-roles", default="")
            command.add_argument("--save-txt", action="store_true")
        elif name == "folders":
            command.add_argument("--sheet", required=True)
            command.add_argument("--upper-col", required=True)
            command.add_argument("--lower-col", default="")
            command.add_argument("--number-mode", choices=("none", "digit", "alpha"), default="digit")
            command.add_argument("--output-dir")
            command.add_argument("--create", action="store_true")
    return cli


def execute(args):
    repo_root = args.repo_root.resolve()
    if not (repo_root / "backend" / "xls_merge.py").is_file():
        raise ValueError("backend/를 찾을 수 없습니다. --repo-root에 Work-Helper 저장소를 지정하세요.")
    sys.path.insert(0, str(repo_root))
    from backend import folders, mail_recipients, xls_merge

    if args.command == "merge":
        if folders.INVALID_CHARS.search(args.prefix):
            raise ValueError("접두어에 파일명 금지 문자를 사용할 수 없습니다.")
        sheets = xls_merge.parse_sheets(args.sheets)
        headers = [int(value.strip()) for value in args.headers.split(",")]
        if not sheets or len(headers) not in (1, len(sheets)) or any(h < 0 for h in headers):
            raise ValueError("시트명과 0 이상의 헤더 행을 지정하세요. 헤더 행은 하나 또는 시트 수만큼 필요합니다.")
        return xls_merge.run(args.folder, args.prefix, args.sheets, args.headers, args.recursive)

    if args.command in ("roster", "recipients"):
        analyzed = mail_recipients.analyze(args.file)
        if "_state" not in analyzed:
            return analyzed
        if args.command == "roster":
            return analyzed["meta"]
        to_roles = mail_recipients.parse_list(args.to_roles)
        cc_roles = mail_recipients.parse_list(args.cc_roles)
        unknown = set(to_roles + cc_roles) - set(analyzed["meta"]["roles"])
        if unknown or not (to_roles or cc_roles):
            raise ValueError("주소록에 있는 수신 또는 참조 역할을 지정하세요. 알 수 없는 역할: " + ", ".join(sorted(unknown)))
        result = mail_recipients.compute(analyzed["_state"], args.teams, to_roles, cc_roles)
        if result.get("ok") and args.save_txt:
            result["out_path"] = mail_recipients.save_txt(analyzed["_state"], result, to_roles, cc_roles)
        return result

    info = folders.load_excel(args.file)
    if args.command == "folder-info":
        return info
    sheet = next((s for s in info["sheets"] if s["name"] == args.sheet), None)
    if sheet is None or sheet["header_row"] is None:
        raise ValueError("시트 또는 헤더를 찾을 수 없습니다. folder-info로 확인하세요.")
    for column in (args.upper_col, args.lower_col):
        if column and column not in sheet["headers"]:
            raise ValueError(f"컬럼을 찾을 수 없습니다: {column}")
    if args.create:
        if not args.output_dir:
            raise ValueError("실제 생성에는 --output-dir이 필요합니다.")
        return folders.run(args.file, args.sheet, args.upper_col, args.lower_col,
                           args.number_mode, args.output_dir)
    return folders.preview(args.file, args.sheet, args.upper_col, args.lower_col, args.number_mode)


def main():
    args = parser().parse_args()
    try:
        result = execute(args)
    except Exception as error:
        result = {"ok": False, "error": str(error)}
    print(json.dumps(result, ensure_ascii=False))
    return 0 if result.get("ok") else 1


if __name__ == "__main__":
    sys.exit(main())
