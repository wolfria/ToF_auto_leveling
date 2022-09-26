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
    # adventure
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

    # challenge
    while True:
        if pyautogui.locateOnScreen(f'./images/{lang}/challenge_active.png' , confidence=0.7) is not None:
            break
        print('Looking for Challenge tab')
        challengePosition = pyautogui.locateOnScreen(f'./images/{lang}/challenge.png' , confidence=0.7)
        if challengePosition is not None:
            print('Challenge tab is found. Click it')
            pyautogui.click(challengePosition)
            time.sleep(1)

    # frontier clash
    while True:
        shell.SendKeys('%')
        win32gui.SetForegroundWindow(tofApp)
        frontierClashPosition = pyautogui.locateOnScreen(f'./images/{lang}/frontier_clash.png' , confidence=0.7)
        print(frontierClashPosition)
        if frontierClashPosition is not None:
            print('Looking for Frontier Crash')
            pyautogui.click(frontierClashPosition.left + 50, frontierClashPosition.top + 420)
            time.sleep(2)
            break
        else:
            print("Can't find Frontier Crash, so I'll look for it")
            trainingPosition = pyautogui.locateOnScreen(f'./images/{lang}/training.png' , confidence=0.7)
            pyautogui.moveTo(trainingPosition)
            print(trainingPosition.left)
            pyautogui.dragTo(trainingPosition.left - 300, trainingPosition.top, duration=0.5, button="left")
            continue

    # join
    print("Participate in Frontier Clash")
    battleZoneGoPosition = pyautogui.locateOnScreen(f'./images/{lang}/frontier_clash_go.png' , confidence=0.7)
    pyautogui.click(battleZoneGoPosition)
    time.sleep(1)

    while pyautogui.locateOnScreen(f'./images/{lang}/jump_device_is_starting.png' , confidence=0.7) is None:
        print('Waiting for matching...')
        # matching
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

        # approve
        approvePosition = pyautogui.locateOnScreen(f'./images/{lang}/approve.png' , confidence=0.7)
        if approvePosition is not None:
            pyautogui.click(approvePosition)

    print('Successfully matched!')

    print('Waiting for the combat to start...')
    time.sleep(25)

    # auto mode on
    while True:
        shell.SendKeys('%')
        win32gui.SetForegroundWindow(tofApp)
        print('Waiting for auto mode to be activated...')
        if pyautogui.locateOnScreen(f'./images/{lang}/auto_on.png' , confidence=0.8) is not None:
            break
        else: 
            pyautogui.keyDown('altleft')
            autoOffPosition = pyautogui.locateOnScreen(f'./images/{lang}/auto_off.png' , confidence=0.7)
            pyautogui.click(autoOffPosition)
            pyautogui.keyUp('altleft')

    # combat to start
    shell.SendKeys('%')
    win32gui.SetForegroundWindow(tofApp)
    pyautogui.keyDown('w')
    time.sleep(2.5)
    pyautogui.keyUp('w')
    pyautogui.keyDown('a')
    time.sleep(3)
    pyautogui.keyUp('a')
    while True:
        if pyautogui.locateOnScreen(f'./images/{lang}/activate.png' , confidence=0.7) is not None:
            pyautogui.press('f', presses=5)
            break
        if pyautogui.locateOnScreen(f'./images/{lang}/synced_to_assistance_system.png' , confidence=0.7) is not None:
            break
        elif pyautogui.locateOnScreen(f'./images/{lang}/first_wave_arrives.png' , confidence=0.7) is not None:
            break
        elif pyautogui.locateOnScreen(f'./images/{lang}/tap_anywhere_to_close.png' , confidence=0.7) is not None:
            print('Challenge failed')
            print('Next...')
            time.sleep(25)
            continue
        shell.SendKeys('%')
        win32gui.SetForegroundWindow(tofApp)
        pyautogui.keyDown('w')
        time.sleep(0.5)
        pyautogui.keyUp('w')
        pyautogui.move(0, 5)

    print('In combat...')
    time.sleep(400)
    combatOverWaitStartTime = time.time()

    # wait combat over
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

    # treasure result
    while True:
        if pyautogui.locateOnScreen(f'./images/{lang}/tap_anywhere_to_close.png' , confidence=0.7) is not None:
            break
        print('Skip Treasure')
        skipPosition = pyautogui.locateOnScreen(f'./images/{lang}/skip.png' , confidence=0.7)
        if skipPosition is not None:
            pyautogui.moveTo(skipPosition.left + 10, skipPosition.top, duration=0.5)
            pyautogui.click()

    # challenge success
    while True:
        if pyautogui.locateOnScreen('./images/commons/exit_combat_result.png' , confidence=0.7) is not None:
            break
        print('Skip Combat Results')
        tapAnyWhereToClosePosition = pyautogui.locateOnScreen(f'./images/{lang}/tap_anywhere_to_close.png' , confidence=0.7)
        if tapAnyWhereToClosePosition is not None:
            pyautogui.moveTo(tapAnyWhereToClosePosition.left, tapAnyWhereToClosePosition.top - 30, duration=0.5)
            pyautogui.click()

    # combat over
    while True:
        if pyautogui.locateOnScreen('./images/commons/exit_combat.png' , confidence=0.7) is not None:
            break
        exitCombatResultPosition = pyautogui.locateOnScreen('./images/commons/exit_combat_result.png' , confidence=0.7)
        if exitCombatResultPosition is not None:
            pyautogui.click(exitCombatResultPosition.left, exitCombatResultPosition.top)

    # exit combat
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
