import json
import pyautogui
import time
from PIL import Image
import numpy as np
import easyocr
import winsound
import os
import keyboard
from difflib import SequenceMatcher
import pandas as pd
import logging
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor

# 配置
CONFIG_FILE = 'keys.json'
LOG_FILE = 'price_log.xlsx'

# EasyOCR 初始化（GPU加速）
reader_eng = easyocr.Reader(['en'], gpu=True)
reader_chn = easyocr.Reader(['ch_sim'], gpu=True)
executor = ThreadPoolExecutor(max_workers=2)

# 日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger()

# 全局变量
keys_config = None
is_running = False
is_paused = False
screen_width, screen_height = pyautogui.size()

# 初始化日志
if not os.path.exists(LOG_FILE):
    df = pd.DataFrame(columns=['Time', 'Card_Name', 'Target_Name', 'Price', 'Purchased'])
    df.to_excel(LOG_FILE, index=False)

def load_keys_config():
    global keys_config
    if keys_config is not None:
        return keys_config
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            config = json.load(f)
            keys_config = config.get('keys', [])
            for card in keys_config:
                card['buy_count'] = 0
            logger.info("配置文件加载成功")
            return keys_config
    except Exception as e:
        logger.error(f"加载配置失败: {str(e)}")
        return []

def take_screenshot(region, threshold):
    screenshot = pyautogui.screenshot(region=region)
    gray_image = screenshot.convert('L')
    binary_image = gray_image.point(lambda p: 255 if p > threshold else 0)
    binary_image = Image.eval(binary_image, lambda x: 255 - x)
    screenshot.close()
    return binary_image

def getCardPrice():
    region = (
        int(screen_width * 0.80),
        int(screen_height * 0.85),
        int(screen_width * 0.09),
        int(screen_height * 0.03)
    )
    image = take_screenshot(region=region, threshold=95)
    image.save("price_screenshot.png")
    result = reader_eng.readtext(np.array(image), detail=0)
    if result:
        text = result[0].replace(",", "").strip()
        try:
            price = int(''.join(filter(str.isdigit, text)))
            logger.info(f"提取的价格文本: {price}")
            return price
        except:
            logger.warning(f"解析价格失败: {text}")
            return None
    return None

def getCardName():
    region = (
        int(screen_width * 0.765),
        int(screen_height * 0.148),
        int(screen_width * 0.1),
        int(screen_height * 0.025)
    )
    screenshot = take_screenshot(region=region, threshold=150)
    screenshot.save("./s.png")
    result = reader_chn.readtext(np.array(screenshot), detail=0)
    if result:
        return result[0].replace(" ", "").strip()
    return ""

def log_to_excel(card_name, target_name, price, purchased):
    current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    new_data = pd.DataFrame({
        'Time': [current_time],
        'Card_Name': [card_name],
        'Target_Name': [target_name],
        'Price': [price if price is not None else 'N/A'],
        'Purchased': [purchased]
    })
    try:
        existing_data = pd.read_excel(LOG_FILE)
        updated_data = pd.concat([existing_data, new_data], ignore_index=True)
        updated_data.to_excel(LOG_FILE, index=False)
    except:
        new_data.to_excel(LOG_FILE, index=False)

def price_check_flow(card_info):
    max_buy_limit = 2
    if card_info.get('buy_count', 0) >= max_buy_limit:
        logger.info(f"{card_info['name']} 已购买 {max_buy_limit} 次，跳过")
        return False

    pyautogui.moveTo(card_info['position'][0]*screen_width, card_info['position'][1]*screen_height)
    pyautogui.click()
    time.sleep(0.02)

    try:
        card_name = getCardName()
        time.sleep(0.02)
        current_price = getCardPrice()
        if current_price is None:
            logger.warning("价格识别失败，跳过")
            log_to_excel(card_name, card_info.get("name"), None, False)
            time.sleep(0.05)
            pyautogui.press('esc')
            return False
    except Exception as e:
        logger.error(f"识别失败: {str(e)}")
        log_to_excel("", card_info.get("name"), None, False)
        time.sleep(0.02)
        pyautogui.press('esc')
        return False

    base_price = card_info.get('base_price', 0)
    ideal_price = card_info.get('ideal_price', base_price)
    max_price = base_price * 1.1
    premium = ((current_price / base_price) - 1) * 100

    similarity = SequenceMatcher(None, card_name, card_info['name']).ratio()
    logger.info(f"门卡识别: {card_name} vs 目标: {card_info['name']} 相似度: {similarity:.2%}")

    if similarity < 0.7:
        logger.info("相似度不足，跳过")
        log_to_excel(card_name, card_info['name'], current_price, False)
        pyautogui.press('esc')
        return False

    purchased = False
    if premium < 0 or current_price < ideal_price or current_price - base_price <= 100:
        if card_info.get('buyMax', 0) == 1:
            pyautogui.moveTo(screen_width*0.9104, screen_height*0.7807)
            pyautogui.click()
            time.sleep(0.02)
        pyautogui.moveTo(screen_width*0.825, screen_height*0.852)
        pyautogui.click()
        logger.info(f"[✓] 购买成功: {card_name} ¥{current_price} 溢价: {premium:.2f}%")
        card_info['buy_count'] += 1
        purchased = True
        time.sleep(0.02)
        pyautogui.press('esc')
    else:
        logger.info("价格过高，已跳过")
        pyautogui.press('esc')

    log_to_excel(card_name, card_info['name'], current_price, purchased)
    return purchased

def start_loop():
    global is_running, is_paused
    is_running = True
    is_paused = False
    logger.info("已启动监控")

def stop_loop():
    global is_running, is_paused
    is_running = False
    is_paused = False
    logger.info("已停止监控")

def all_cards_completed(cards):
    return all(card.get('buy_count', 0) >= 2 for card in cards)

def main():
    global is_running, is_paused
    keys_config = load_keys_config()
    if not keys_config:
        logger.error("配置文件加载失败")
        return

    cards_to_monitor = [card for card in keys_config if card.get('wantBuy', 0) == 1]
    if not cards_to_monitor:
        logger.error("没有需要监控的门卡")
        return

    for card in cards_to_monitor:
        logger.info(f"准备监控: {card['name']}")

    keyboard.add_hotkey('f8', start_loop)
    keyboard.add_hotkey('f9', stop_loop)

    logger.info("按 F8 开始监控，F9 停止")

    while True:
        if is_running and not is_paused:
            futures = [executor.submit(price_check_flow, card) for card in cards_to_monitor]
            for f in futures: f.result()

            if all_cards_completed(cards_to_monitor):
                logger.info("所有门卡已完成购买")
                is_running = False
        elif is_paused:
            logger.info("暂停中...")
            time.sleep(1)
        else:
            time.sleep(0.05)

if __name__ == "__main__":
    main()
