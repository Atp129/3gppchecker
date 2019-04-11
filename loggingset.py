#!/usr/bin/env python
# encoding: utf-8
# # @author: Hongkang LI
# @license: (C) Copyright 1990-2021, UNISOC Technologies Corporation Limited.
# @contact: romain.li@unisoc.com
# @software: unisoc
# @file: loggingset.py
# @time: 2019/4/11 10:31
# @desc: make the default log setting, peck the logging information
#

import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


def info(string, *args):
    logger.info(string, *args)


def debug(string, *args):
    logger.debug(string, *args)


def warning(string, *args):
    logger.warning(string, *args)