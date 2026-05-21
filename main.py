import win32com.client
import datetime
import os
import time


def download_and_merge_ppts():
    base_dir = os.getcwd()
    download_dir = os.path.join(base_dir, "PPT_다운로드")

    if not os.path.exists(download_dir):
        os.makedirs(download_dir)

    # 취합 대상 부서 및 담당자 명단
    target_teams = {
        "허용민": {"team": "연구기획팀", "rank": 1, "submitted_files": []},
        "김태원": {"team": "심혈관팀", "rank": 2, "submitted_files": []},
        "한예지": {"team": "급성감염팀", "rank": 3, "submitted_files": []},
        "정진용": {"team": "Cancer팀", "rank": 4, "submitted_files": []},
        "이소희": {"team": "호르몬팀", "rank": 5, "submitted_files": []},
        "김영은": {"team": "치료용항체팀", "rank": 6, "submitted_files": []},
        "김세희": {"team": "갑상선팀", "rank": 7, "submitted_files": []},
        "함은선": {"team": "당뇨팀", "rank": 8, "submitted_files": []}
    }

    downloaded_info = []

    # ★ [핵심] 기준일 설정: 코드를 실행하는 "오늘(작업하는 날짜)"
    target_date = datetime.date.today()

    # ==========================================
    # STEP 1: 아웃룩에서 메일 검색 및 첨부파일 다운로드
    # ==========================================
    print("=" * 70)
    print(f"[ STEP 1. 이메일 점검 및 첨부파일 다운로드 작업 ] - 기준일: {target_date}")
    print("=" * 70)

    try:
        outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        inbox = namespace.GetDefaultFolder(6)

        messages = inbox.Items
        messages.Sort("[ReceivedTime]", True)

        found_files_count = 0

        for message in messages:
            if message.Class == 43:
                # 메일이 도착한 날짜 추출
                msg_date = datetime.date(message.ReceivedTime.year, message.ReceivedTime.month,
                                         message.ReceivedTime.day)

                # 1. 작업하는 날짜(오늘)보다 과거의 메일이면 탐색을 즉시 멈춤
                if msg_date < target_date:
                    break

                    # 2. 작업하는 날짜(오늘)에 확인된 메일만 작업 진행
                if msg_date == target_date:
                    for attachment in message.Attachments:
                        file_name = attachment.FileName
                        file_name_lower = file_name.lower()

                        # 3. 파일명에 날짜가 있는지 상관없이, 확장자가 PPT이고 '회의자료'라는 글자만 있으면 OK
                        if file_name_lower.endswith(('.ppt', '.pptx')) and '회의자료' in file_name:
                            found_files_count += 1

                            # 중복 덮어쓰기 방지를 위해 다운로드할 파일명 앞에 수신시간만 살짝 붙여줍니다.
                            time_str = message.ReceivedTime.strftime('%H%M%S')
                            safe_file_name = f"{time_str}_{file_name}"

                            save_path = os.path.join(download_dir, safe_file_name)
                            abs_save_path = os.path.abspath(save_path)

                            attachment.SaveAsFile(abs_save_path)

                            sender_name = message.SenderName
                            rank = 99
                            team_name = "기타(타부서)"

                            for name, info in target_teams.items():
                                if name in sender_name:
                                    rank = info["rank"]
                                    team_name = info["team"]
                                    info["submitted_files"].append(file_name)
                                    break

                            downloaded_info.append({
                                "path": abs_save_path,
                                "sender": sender_name,
                                "team": team_name,
                                "rank": rank,
                                "file_name": safe_file_name
                            })

                            print(f" 📥 [다운로드 완료] 보낸사람 : {sender_name} ({team_name})")
                            print(f"    - 원본 파일명 : {file_name}")
                            print(f"    - 저장 파일명 : {safe_file_name}")
                            print(f"    - 취합 순위   : {rank if rank != 99 else '미지정 (맨 끝)'}")
                            print("-" * 70)

        if found_files_count == 0:
            print(f" ⚠️ {target_date} 오늘 도착한 메일 중 조건에 맞는 회의자료(PPT)가 없습니다.")

    except Exception as e:
        print(f"아웃룩 메일 처리 중 오류가 발생했습니다: {e}")
        return

    # ==========================================
    # STEP 2: 부서별 자료 제출 및 점검 현황 리포트
    # ==========================================
    print("\n" + "=" * 70)
    print("[ STEP 2. 부서별 자료 제출 점검 현황 리포트 ]")
    print("=" * 70)

    submitted_count = 0
    missing_count = 0

    for name, info in target_teams.items():
        team_str = f"{info['team']} ({name})"
        files = info["submitted_files"]

        if files:
            submitted_count += 1
            print(f" ✅ [제출] {team_str:<15} -> 📎 원본 파일명: {', '.join(files)}")
        else:
            missing_count += 1
            print(f" ❌ [미제출] {team_str}")

    print("-" * 70)
    print(f" 📊 총계 : 제출 {submitted_count}팀 / 미제출 {missing_count}팀")
    print("=" * 70)

    # ==========================================
    # STEP 3: 순서에 맞게 정렬 후 파워포인트 통합
    # ==========================================
    if not downloaded_info:
        print("\n ⚠️ 병합할 파일이 없어 작업을 종료합니다.")
        return

    downloaded_info.sort(key=lambda x: (x["rank"], x["file_name"]))

    print("\n" + "=" * 70)
    print("[ STEP 3. 파워포인트 파일 순차 병합 작업 ]")
    print("=" * 70)

    try:
        ppt_app = win32com.client.gencache.EnsureDispatch("PowerPoint.Application")
        ppt_app.Visible = 1

        merged_ppt = ppt_app.Presentations.Add()

        total_files = len(downloaded_info)
        for idx, info in enumerate(downloaded_info):
            file_path = info["path"]
            merged_ppt.Slides.InsertFromFile(file_path, merged_ppt.Slides.Count)

            print(f" 🔄 [{idx + 1}/{total_files}] 병합 중... {info['team']} ({info['sender']})")
            print(f"    -> 대상 파일: {info['file_name']}")

        merged_file_path = os.path.abspath(os.path.join(base_dir, f"통합_회의자료_{target_date.strftime('%Y%m%d')}.pptx"))

        if os.path.exists(merged_file_path):
            os.remove(merged_file_path)

        time.sleep(2)
        merged_ppt.SaveAs(merged_file_path, 24)

        merged_ppt.Close()
        ppt_app.Quit()

        print("-" * 70)
        print(" 🎉 파일 병합 작업이 성공적으로 완료되었습니다!")
        print(f" 📁 최종 저장 위치: {merged_file_path}")
        print("=" * 70)

    except Exception as e:
        print(f"파워포인트 파일 통합 중 오류가 발생했습니다: {e}")


if __name__ == "__main__":
    download_and_merge_ppts()
