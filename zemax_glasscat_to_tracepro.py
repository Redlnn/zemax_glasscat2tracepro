"""
Zemax 玻璃库导入 TracePro

仅在 TracePro 7.x 测试，不一定适用于其他版本
运行前请至少启动 TracePro 一次

GitHub: https://github.com/Redlnn/zemax_glasscat2tracepro/
"""

import logging
import os
import re
import sqlite3
import sys
import tkinter as tk
import winreg
from pathlib import Path
from tkinter import filedialog
from typing import TypedDict

# 折射率公式对应关系
# zemax to tracepro
equation_map = {
    "1": "1",  # Schoot
    "2": "2",  # Sellmeier 1
    "3": "4",  # Herzberger
}

root = tk.Tk()
root.withdraw()

path = filedialog.askopenfilename(
    initialdir=Path(
        winreg.QueryValueEx(
            winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders",
            ),
            "Personal",
        )[0]
    ),
    title="选择玻璃库",
    filetypes=(("Zemax 玻璃库文件", "*.agf"),),
)

if not path:
    print('未选择玻璃库')
    sys.exit(1)

glasscat_path = Path(path)

if not glasscat_path.exists():
    print('所选玻璃库不存在')
    sys.exit(1)

appdata = os.getenv("APPDATA")

if appdata is None:
    print('系统环境变量 %APPDATA% 异常，请修复或重装系统')
    sys.exit(1)

database_path = (
    Path(appdata) / 'Lambda Research Corporation' / 'TracePro' / 'TracePro.db'
)

if not database_path.exists():
    print('TracePro 数据库不存在，请至少启动 TracePro 一次')

logging.basicConfig(
    level=logging.DEBUG, format='%(asctime)s | %(levelname)s | %(message)s'
)

glasscat_name = glasscat_path.name.rstrip(glasscat_path.suffix).replace("-", "_")

if not re.match("^[a-zA-Z0-9_]+$", glasscat_name):
    raise ValueError("文件名不合法")
logging.info(f'玻璃库名称: {glasscat_name}')

logging.info(f'打开TracePro数据库: {database_path}')
conn = sqlite3.connect(database_path)
cur = conn.cursor()

logging.info(f'创建数据库表: MATL-{glasscat_name}')
cur.execute(
    (
        f"CREATE TABLE IF NOT EXISTS 'MATL-{glasscat_name}' "
        "( 'ID' INTEGER PRIMARY KEY, 'Name' TEXT, 'Description' TEXT, 'MILSPEC' TEXT, "
        "'WavelengthStart' REAL, 'WavelengthEnd' REAL, 'UserData' INTEGER, 'Type' "
        "INTEGER, 'NTerm' REAL, 'Term0' REAL, 'Term1' REAL, 'Term2' REAL, 'Term3' "
        "REAL, 'Term4' REAL, 'Term5' REAL, 'Term6' REAL, 'Term7' REAL, 'Term8' REAL, "
        "'Term9' REAL)"
    )
)
logging.info(f'创建数据库表: MATL-{glasscat_name}-Data')
cur.execute(
    (
        f"CREATE TABLE IF NOT EXISTS 'MATL-{glasscat_name}-Data' ( 'ID' INTEGER "
        "PRIMARY KEY, 'Name' TEXT, 'Temperature' REAL, 'Wavelength' REAL, 'Index-real' "
        "REAL, 'Index-imag' REAL)"
    )
)
logging.info(f'创建数据库表: MATL-{glasscat_name}-UniaxialData')
cur.execute(
    (
        f"CREATE TABLE IF NOT EXISTS 'MATL-{glasscat_name}-UniaxialData' ( 'ID' "
        "INTEGER PRIMARY KEY, 'Name' TEXT, 'Temperature' REAL, 'Wavelength' REAL, "
        "'Index-real-or' REAL, 'Index-real-ex' REAL, 'Index-imag-or' REAL, "
        "'Index-imag-ex' REAL, 'Gyro-tensor-11' REAL, 'Gyro-tensor-12' REAL, "
        "'Gyro-tensor-13' REAL, 'Gyro-tensor-22' REAL, 'Gyro-tensor-23' REAL, "
        "'Gyro-tensor-33' REAL)"
    )
)

logging.info('创建玻璃库')
res = cur.execute(
    f'SELECT Name FROM "main"."MaterialCatalogs" WHERE Name=\'{glasscat_name}\''
).fetchall()

if not res:
    cur.execute(
        'INSERT INTO "main"."MaterialCatalogs" ("Name", "Table", "Type") VALUES '
        f"('{glasscat_name}', '{glasscat_name}', '0');"
    )

logging.info(f'打开玻璃库: {glasscat_path}')
f = open(glasscat_path, encoding="UTF-8")
try:
    f.readline()
except UnicodeDecodeError:
    f.close()
    f = open(glasscat_path, encoding="UTF-16LE")
else:
    f.seek(0, 0)


class Glass(TypedDict, total=False):
    name: str
    Description: str
    Equation: str
    WaveStart: float
    WaveEnd: float
    coefficient: list[float]


try:
    glass = Glass(name="", Description="")
    res = cur.execute(f'SELECT Name FROM "main"."MATL-{glasscat_name}"').fetchall()
    while line := f.readline():
        lines = line.split()
        if line.startswith("NM"):
            if (
                (glass["name"] not in res)
                and ((glass["name"],) not in res)
                and glass["name"] != ""
            ):
                cur.execute(
                    (
                        'INSERT INTO "main"."MATL-{glasscat_name}" ("Name", '
                        '"Description", "MILSPEC", "WavelengthStart", "WavelengthEnd", '
                        '"UserData", "Type", "NTerm", "Term0", "Term1", "Term2", '
                        '"Term3", "Term4", "Term5", "Term6", "Term7", "Term8", '
                        '"Term9") VALUES (\'{name}\', \'{description}\', \'\', '
                        "'{wav_start}', '{wav_end}', '0', '{equ_type}', "
                        "'0.0', '{term1}', '{term2}', '{term3}', '{term4}', "
                        "'{term5}', '{term6}', '0.0', '0.0', '0.0', '0.0');"
                    ).format(
                        glasscat_name=glasscat_name,
                        name=glass["name"],
                        description=glass["Description"],
                        wav_start=glass["WaveStart"],
                        wav_end=glass["WaveEnd"],
                        equ_type=glass["Equation"],
                        term1=glass["coefficient"][0],
                        term2=glass["coefficient"][1],
                        term3=glass["coefficient"][2],
                        term4=glass["coefficient"][3],
                        term5=glass["coefficient"][4],
                        term6=glass["coefficient"][5],
                    )
                )
                logging.info(
                    (
                        f'玻璃名称: {glass["name"]}, 描述: {glass["Description"]}'
                        f', 波长范围: {glass["WaveStart"]}~{glass["WaveEnd"]}'
                    )
                )

                glass = Glass(name="", Description="")
            glass["name"] = lines[1]
            glass["Equation"] = equation_map[str(int(float(lines[2])))]
        if line.startswith("GC"):
            glass["Description"] = line.split(" ", 1)[1].rstrip()
        if line.startswith("CD"):
            if glass["Equation"] in ("1", "4"):
                glass["coefficient"] = [
                    float(lines[1]),
                    float(lines[2]),
                    float(lines[3]),
                    float(lines[4]),
                    float(lines[5]),
                    float(lines[6]),
                ]
            elif glass["Equation"] == "2":
                glass["coefficient"] = [
                    float(lines[1]),
                    float(lines[3]),
                    float(lines[5]),
                    float(lines[2]),
                    float(lines[4]),
                    float(lines[6]),
                ]
            else:
                raise ValueError("Unknow Equation")
        if line.startswith("LD"):
            glass["WaveStart"] = float(lines[1])
            glass["WaveEnd"] = float(lines[2])
    if (
        (glass["name"] not in res)
        and ((glass["name"],) not in res)
        and glass["name"] != ""
    ):
        cur.execute(
            (
                'INSERT INTO "main"."MATL-{glasscat_name}" ("Name", "Description", '
                '"MILSPEC", "WavelengthStart", "WavelengthEnd", "UserData", "Type", '
                '"NTerm", "Term0", "Term1", "Term2", "Term3", "Term4", "Term5", "Term6"'
                ', "Term7", "Term8", "Term9") VALUES (\'{name}\', \'{description}\', '
                "'', '{wav_start}', '{wav_end}', '0', '{equ_type}', '0.0', "
                "'{term1}', '{term2}', '{term3}', '{term4}', '{term5}', '{term6}',"
                " '0.0', '0.0', '0.0', '0.0');"
            ).format(
                glasscat_name=glasscat_name,
                name=glass["name"],
                description=glass["Description"],
                wav_start=glass["WaveStart"],
                wav_end=glass["WaveEnd"],
                equ_type=glass["Equation"],
                term1=glass["coefficient"][0],
                term2=glass["coefficient"][1],
                term3=glass["coefficient"][2],
                term4=glass["coefficient"][3],
                term5=glass["coefficient"][4],
                term6=glass["coefficient"][5],
            )
        )
        logging.info(
            (
                f'玻璃名称: {glass["name"]}, 描述: {glass["Description"]}, '
                f'波长范围: {glass["WaveStart"]}~{glass["WaveEnd"]}'
            )
        )
except Exception:
    cur.close()
    conn.rollback()
    logging.critical('出错了，已撤回数据库操作')
    raise
else:
    cur.close()
    conn.commit()
    logging.info('已保存数据库')
finally:
    conn.close()
    f.close()
