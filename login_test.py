#!/usr/bin/python
# -*- coding: utf-8 -*-
# Copyright 2023, 2026 Jack Consoli.  All rights reserved.
#
# NOT BROADCOM SUPPORTED
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may also obtain a copy of the License at
# https://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.
"""
:mod:`login_test` - Login and logout. Used to validate the HTTP connections

Version Control::

    +-----------+---------------+-----------------------------------------------------------------------------------+
    | Version   | Last Edit     | Description                                                                       |
    +===========+===============+===================================================================================+
    | 4.0.0     | 04 Aug 2023   | Re-Launch                                                                         |
    +-----------+---------------+-----------------------------------------------------------------------------------+
    | 4.0.1     | 20 Feb 2026   | Updated copyright notice.                                                         |
    +-----------+---------------+-----------------------------------------------------------------------------------+
"""

__author__ = 'Jack Consoli'
__copyright__ = 'Copyright 2023 Consoli Solutions, LLC'
__date__ = '20 Feb 2026'
__license__ = 'Apache License, Version 2.0'
__email__ = 'jack_consoli@yahoo.com'
__maintainer__ = 'Jack Consoli'
__status__ = 'Released'
__version__ = '4.0.1'

import argparse
import brcdapi.log as brcdapi_log
import brcdapi.util as brcdapi_util
import brcdapi.sannav_auth as sannav_auth
import brcddb.brcddb_common as brcddb_common

_DOC_STRING = False  # Should always be False. Prohibits any code execution. Only useful for building documentation
_DEBUG = False  # When True, use _DEBUG_IP, _DEBUG_ID, _DEBUG_PW, AND _DEBUG_OUTF instead of passed arguments
_DEBUG_IP = 'xx.xxx.x.xxx'
_DEBUG_ID = 'Administrator'
_DEBUG_PW = 'password'
_DEBUG_SEC = 'self'
_DEBUG_LOG = '_logs'
_DEBUG_NL = True


def parse_args():
    """Parses the module load command line

    :return ip_addr: IP address
    :rtype ip_addr: str
    :return id: User ID
    :rtype id: str
    :return pw: Password
    :rtype pw: str
    :return http_sec: Type of HTTP security
    :rtype http_sec: str
    """
    global _DEBUG_IP, _DEBUG_ID, _DEBUG_PW, _DEBUG_SEC, _DEBUG_LOG, _DEBUG_NL

    if _DEBUG:
        return _DEBUG_IP, _DEBUG_ID, _DEBUG_PW, _DEBUG_SEC, _DEBUG_LOG, _DEBUG_NL

    parser = argparse.ArgumentParser(description='Capture (GET) requests from a chassis')
    parser.add_argument('-ip', help='IP address', required=True)
    parser.add_argument('-id', help='User ID', required=True)
    parser.add_argument('-pw', help='Password', required=True)
    parser.add_argument('-s', help='\'CA\' or \'self\' for HTTPS mode.', required=False)
    buf = '(Optional) Directory where log file is to be created. Default is to use the current directory. The log '\
          'file name will always be "Log_xxxx" where xxxx is a time and date stamp.'
    parser.add_argument('-log', help=buf, required=False,)
    buf = '(Optional) No parameters. When set, a log file is not created. The default is to create a log file.'
    parser.add_argument('-nl', help=buf, action='store_true', required=False)
    args = parser.parse_args()
    return args.ip, args.id, args.pw, args.s, args.log, args.nl


def psuedo_main():
    """Basically the main(). Did it this way so it can easily be used as a standalone module or called from another.

    :return: Exit code. See exist codes in brcddb.brcddb_common
    :rtype: int
    """
    # Get user input
    ip, user_id, pw, sec, log, nl = parse_args()
    if sec is None:
        sec = 'none'
    ml = ['WARNING!!! Debug is enabled'] if _DEBUG else list()
    ml.append('IP:          ' + brcdapi_util.mask_ip_addr(ip, True))
    ml.append('ID:          ' + user_id)
    ml.append('security:    ' + sec)

    # Open the log file
    brcdapi_log.open_log(log)
    brcdapi_log.log(ml, True)

    # Login
    brcdapi_log.log('Attempting login', True)
    session = sannav_auth.login(user_id, pw, ip, sec)
    if sannav_auth.is_error(session):
        brcdapi_log.log('Login failed. Error message is:', True)
        brcdapi_log.log(sannav_auth.formatted_error_msg(session), True)
        return brcddb_common.EXIT_STATUS_API_ERROR

    # Logout
    brcdapi_log.log('Login succeeded. Attempting logout', True)
    obj = sannav_auth.logout(session)
    if sannav_auth.is_error(obj):
        brcdapi_log.log('Logout failed. Error message is:', True)
        brcdapi_log.log(sannav_auth.formatted_error_msg(obj), True)
        return -1
    brcdapi_log.log('Logout succeeded.', True)

    # Close the log file
    brcdapi_log.close_log('')
    return brcddb_common.EXIT_STATUS_OK


###################################################################
#
#                    Main Entry Point
#
###################################################################
_ec = 0
if _DOC_STRING:
    print('_DOC_STRING is True. No processing')
else:
    _ec = psuedo_main()
exit(_ec)
