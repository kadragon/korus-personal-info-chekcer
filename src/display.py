"""
이 모듈은 rich 라이브러리를 사용하여 터미널에 스타일링된 출력을 처리합니다.
헤더, 결과, 요약, 오류 메시지 등을 출력하기 위한 중앙 집중식 함수를 제공합니다.
"""

from rich.console import Console
from rich.panel import Panel
from rich.text import Text

# 전역 콘솔 객체 생성
console = Console(markup=True)


def print_header(title: str):
    """프로그램의 메인 헤더를 출력합니다."""
    console.rule(f"[bold cyan]{title}[/bold cyan]", style="cyan")


def print_checker_header(title: str):
    """각 검사기의 헤더를 출력합니다."""
    console.print(Text(f"\n--- {title} 검사 ---", style="bold yellow"), justify="left")


def print_result(is_detected: bool, description: str, filename: str | None = None):
    """
    검사 결과를 [양호] 또는 [검출] 형식으로 출력합니다.
    검출된 경우, 관련 파일명도 함께 출력합니다.
    """
    if is_detected:
        status = Text("[검출]", style="bold red")
        message = f"{description} -> 저장 파일: {filename}"
    else:
        status = Text("[양호]", style="bold green")
        message = description

    console.print(f"  {status} {message}")


def print_summary(folder_path: str, total_count: int | None = None):
    """모든 작업 완료 후 최종 요약 메시지를 출력합니다."""

    summary_text = "[bold green]모든 작업이 완료되었습니다.[/bold green]\n\n"

    if total_count is not None:
        summary_text += f"붙임2, 3, 4 원본 데이터 건수 합계: {total_count}건\n\n"

    summary_text += "최종 결과는 아래 폴더를 확인해주세요.\n"
    summary_text += f"[underline blue]{folder_path}[/underline blue]"

    console.print(
        Panel(
            summary_text,
            title="[bold]작업 완료[/bold]",
            border_style="green",
        )
    )


def print_error(message: str):
    """오류 메시지를 출력합니다."""
    console.print(f"[bold red]오류:[/bold red] {message}")


def print_info(message: str):
    """단순 정보 메시지를 출력합니다."""
    console.print(f"[cyan]▶[/cyan] {message}")


def print_zip_header():
    """압축 작업 헤더를 출력합니다."""
    console.print(Text("\n--- 압축 작업 시작 ---", style="bold yellow"), justify="left")


def print_zip_result(zip_name: str, num_files: int):
    """압축 작업 결과를 출력합니다."""
    console.print(
        f"  [bold green]✅[/bold green] {zip_name} 생성 ({num_files}개 파일 포함)"
    )


def print_zip_warning(prefix: str):
    """압축할 파일이 없을 때 경고를 출력합니다."""
    console.print(f"  [yellow]⚠️[/yellow] {prefix}로 시작하는 파일 없음")
