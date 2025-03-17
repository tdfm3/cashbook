#! python3
# -*- coding: utf-8 -*-
# -----------------------------------------------------------------------------
# test_cashbook.py
# テストコード
# :e ++enc=utf8
# -----------------------------------------------------------------------------

import os
import sys
from logging import ERROR, INFO, Formatter, Logger, StreamHandler, getLogger
from pathlib import Path
from typing import Generator

import pytest

from cashbook import disp_useage


def create_logger() -> Logger:
    logger: Logger = getLogger(__name__)

    # NOTE:
    # 重複登録を止める
    for h in logger.handlers[:]:
        logger.removeHandler(h)
        h.close()

    logger.setLevel(INFO)

    format: str = '%(asctime)s | %(levelname)s | %(name)s - %(message)s'

    s_handler: StreamHandler = StreamHandler()
    s_handler.setFormatter(Formatter(format))
    s_handler.setLevel(INFO)

    logger.addHandler(s_handler)

    return logger


@pytest.fixture
def setup() -> Generator[str, None, None]:
    # setup
    msg: str = '現金出納帳を集計するツールです'
    print('!!setup')

    # test
    yield msg

    # teardown
    print('!!done')


def test_disp_useage(setup) -> None:
    msg: str = setup
    sut: str = disp_useage()

    suts: list[str] = sut.splitlines()
    assert suts[0] == msg


def test_base_path() -> None:
    logger: Logger = create_logger()

    # pyinstallerでEXE化する場合は、カレントディレクトリの取得方法に注意
    print('\n')

    # exeから起動する場合
    # base_path: str = os.path.dirname(os.path.abspath(sys.argv[0]))
    base_path: str = os.getcwd()
    ini_path: str = os.path.join(base_path, 'AKBuild.ini')
    logger.info(f"設定ファイル：{ini_path}")

    # fmt: off
    logger.info(f"1. __file__ =>                                      {__file__}")
    logger.info(f"2. os.getcwd() =>                                   {os.getcwd()}")
    logger.info(f"3. os.path.abspath(__file__) =>                     {os.path.abspath(__file__)}")
    logger.info(f"4. Path(__file__).parent =>                         {Path(__file__).parent}")
    logger.info(f"5. sys.executable =>                                {sys.executable}")
    logger.info(f"6. os.path.abspath(sys.argv[0]) =>                  {os.path.abspath(sys.argv[0])}")
    logger.info(f"7. os.path.dirname(os.path.abspath(sys.argv[0])) => {os.path.dirname(os.path.abspath(sys.argv[0]))}")
    # fmt: on


# vim:fenc=utf-8 bomb tw=0:
