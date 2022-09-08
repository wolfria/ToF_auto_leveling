import time

import pyautogui
import win32com.client
import win32gui

lang = 'jp'

tofApp = win32gui.FindWindow(None, 'Tower of Fantasy  ')
shell = win32com.client.Dispatch("WScript.Shell")

shell.SendKeys('%')
win32gui.SetForegroundWindow(tofApp)
time.sleep(1)

print('Starts auto-leveling!')

while True:
    # 探検
    while True:
        if pyautogui.locateOnScreen(f'./images/en/adventure.png' , confidence=0.5) is not None:
            lang = 'en'
        if pyautogui.locateOnScreen(f'./images/{lang}/challenge.png' , confidence=0.5) is not None:
            break
        print('Looking for Adventure icon')
        adventurePosition = pyautogui.locateOnScreen('./images/commons/adventure.png' , confidence=0.5)
        if adventurePosition is not None:
            print('Adventure icon is found. Click it')
            with pyautogui.hold('altleft'):
                pyautogui.moveTo(adventurePosition)
                pyautogui.click(adventurePosition)
            time.sleep(1)

    # 挑戦
    while True:
        if pyautogui.locateOnScreen(f'./images/{lang}/challenge_active.png' , confidence=0.7) is not None:
            break
        print('Looking for Challenge tab')
        challengePosition = pyautogui.locateOnScreen(f'./images/{lang}/challenge.png' , confidence=0.7)
        if challengePosition is not None:
            print('Challenge tab is found. Click it')
            pyautogui.click(challengePosition)
            time.sleep(1)

    # 境界戦闘地帯
    while True:
        shell.SendKeys('%')
        win32gui.SetForegroundWindow(tofApp)
        battleZonePosition = pyautogui.locateOnScreen(f'./images/{lang}/frontier_clash.png' , confidence=0.7)
        if battleZonePosition is not None:
            print('Looking for Frontier Crash')
            pyautogui.click(battleZonePosition.left + 50, battleZonePosition.top + 420)
            time.sleep(2)
            break
        else:
            print("Can't find Frontier Crash, so I'll look for it")
            bygonePhantasmPosition = pyautogui.locateOnScreen(f'./images/{lang}/bygone_phantasm.png' , confidence=0.7)
            trainingPosition = pyautogui.locateOnScreen(f'./images/{lang}/training.png' , confidence=0.7)
            pyautogui.moveTo(bygonePhantasmPosition)
            pyautogui.dragTo(trainingPosition, duration=0.5, button="left")
            continue

    # 参加
    print("Participate in Frontier Clash")
    battleZoneGoPosition = pyautogui.locateOnScreen(f'./images/{lang}/frontier_clash_go.png' , confidence=0.7)
    pyautogui.click(battleZoneGoPosition)
    time.sleep(1)

    while pyautogui.locateOnScreen(f'./images/{lang}/jump_device_is_starting.png' , confidence=0.7) is None:
        print('Waiting for matching...')
        # マッチング
        matchPosition = pyautogui.locateOnScreen(f'./images/{lang}/match.png' , confidence=0.7)
        if matchPosition is not None:
            pyautogui.click(matchPosition)

        while True:
            if pyautogui.locateOnScreen(f'./images/{lang}/approve.png' , confidence=0.7) is not None:
                break
            elif pyautogui.locateOnScreen(f'./images/{lang}/jump_device_is_starting.png' , confidence=0.7) is not None:
                break
            elif pyautogui.locateOnScreen(f'./images/{lang}/match.png' , confidence=0.7) is not None:
                break
            time.sleep(1)

        # 同意
        approvePosition = pyautogui.locateOnScreen(f'./images/{lang}/approve.png' , confidence=0.7)
        if approvePosition is not None:
            pyautogui.click(approvePosition)

    print('Successfully matched!')

    print('Waiting for the combat to start...')
    time.sleep(25)

    while True:
        shell.SendKeys('%')
        win32gui.SetForegroundWindow(tofApp)
        print('Waiting for the combat to start...')
        if pyautogui.locateOnScreen(f'./images/{lang}/please_turn_on_mechanism.png' , confidence=0.7) is not None:
            break
        elif pyautogui.locateOnScreen(f'./images/{lang}/synced_to_assistance_system.png' , confidence=0.7) is not None:
            break
        elif pyautogui.locateOnScreen(f'./images/{lang}/first_wave_arrives.png' , confidence=0.7) is not None:
            break
        elif pyautogui.locateOnScreen(f'./images/{lang}/tap_anywhere_to_close.png' , confidence=0.7) is not None:
            print('Challenge failed')
            time.sleep(25)
            continue
        time.sleep(0.7)

    # オートモードオン
    while True:
        shell.SendKeys('%')
        win32gui.SetForegroundWindow(tofApp)
        print('Waiting for auto mode to be activated...')
        if pyautogui.locateOnScreen(f'./images/{lang}/auto_on.png' , confidence=0.7) is not None:
            break
        else: 
            pyautogui.keyDown('altleft')
            autoOffPosition = pyautogui.locateOnScreen(f'./images/{lang}/auto_off.png' , confidence=0.7)
            pyautogui.click(autoOffPosition)
            pyautogui.keyUp('altleft')

    print('In combat...')
    time.sleep(420)
    combatOverWaitStartTime = time.time()

    # 戦闘終了待ち
    while True:
        shell.SendKeys('%')
        win32gui.SetForegroundWindow(tofApp)
        print('In combat...')
        if pyautogui.locateOnScreen(f'./images/{lang}/skip.png' , confidence=0.7) is not None:
            print('Combat is over!')
            break
        if 240 < (time.time() - combatOverWaitStartTime):
            continue
        time.sleep(5)

    # 挑戦成功
    while True:
        if pyautogui.locateOnScreen(f'./images/{lang}/tap_anywhere_to_close.png' , confidence=0.7) is not None:
            break
        print('Skip Combat Results')
        skipPosition = pyautogui.locateOnScreen(f'./images/{lang}/skip.png' , confidence=0.7)
        if skipPosition is not None:
            pyautogui.moveTo(skipPosition.left + 10, skipPosition.top, duration=0.5)
            pyautogui.click()

    # 戦闘結果
    while True:
        if pyautogui.locateOnScreen('./images/commons/exit_result.png' , confidence=0.7) is not None:
            break
        tapAnywhereToClosePosition = pyautogui.locateOnScreen(f'./images/{lang}/tap_anywhere_to_close.png' , confidence=0.7)
        if tapAnywhereToClosePosition is not None:
            pyautogui.click(tapAnywhereToClosePosition.left, tapAnywhereToClosePosition.top - 50)

    # 戦闘結果閉じる
    while True:
        if pyautogui.locateOnScreen('./images/commons/exit_combat.png' , confidence=0.7) is not None:
            break
        exitResultPosition = pyautogui.locateOnScreen('./images/commons/exit_result.png' , confidence=0.7)
        pyautogui.click(exitResultPosition)

    # 退出
    while True:
        if pyautogui.locateOnScreen(f'./images/{lang}/ok.png' , confidence=0.7) is not None:
            break
        print('Exit from Frontier Crash')
        pyautogui.keyDown('altleft')
        exitBattlePosition = pyautogui.locateOnScreen('./images/commons/exit_combat.png' , confidence=0.7)
        pyautogui.click(exitBattlePosition)
        pyautogui.keyUp('altleft')

    okPosition = pyautogui.locateOnScreen(f'./images/{lang}/ok.png' , confidence=0.7)
    pyautogui.click(okPosition)
    print('Next...')
    time.sleep(15)
    continue
