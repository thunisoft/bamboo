/*
 *
 * (c) Copyright Ascensio System SIA 2010-2019
 *
 * This program is a free software product. You can redistribute it and/or
 * modify it under the terms of the GNU Affero General Public License (AGPL)
 * version 3 as published by the Free Software Foundation. In accordance with
 * Section 7(a) of the GNU AGPL its Section 15 shall be amended to the effect
 * that Ascensio System SIA expressly excludes the warranty of non-infringement
 * of any third-party rights.
 *
 * This program is distributed WITHOUT ANY WARRANTY; without even the implied
 * warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR  PURPOSE. For
 * details, see the GNU AGPL at: http://www.gnu.org/licenses/agpl-3.0.html
 *
 * You can contact Ascensio System SIA at 20A-12 Ernesta Birznieka-Upisha
 * street, Riga, Latvia, EU, LV-1050.
 *
 * The  interactive user interfaces in modified source and object code versions
 * of the Program must display Appropriate Legal Notices, as required under
 * Section 5 of the GNU AGPL version 3.
 *
 * Pursuant to Section 7(b) of the License you must retain the original Product
 * logo when distributing the program. Pursuant to Section 7(e) we decline to
 * grant you any rights under trademark law for use of our trademarks.
 *
 * All the Product's GUI elements, including illustrations and icon sets, as
 * well as technical writing content are licensed under the terms of the
 * Creative Commons Attribution-ShareAlike 4.0 International. See the License
 * terms at http://creativecommons.org/licenses/by-sa/4.0/legalcode
 *
*/

if (Common === undefined) {
    var Common = {};
}

    Common.Gateway = new(function() {
        var me = this,
            $me = $(me);

        var commandMap = {
            'init': function(data) {
                $me.trigger('init', data);
            },

            'openDocument': function(data) {
                $me.trigger('opendocument', data);
            },

            'showMessage': function(data) {
                $me.trigger('showmessage', data);
            },

            'applyEditRights': function(data) {
                $me.trigger('applyeditrights', data);
            },

            'processSaveResult': function(data) {
                $me.trigger('processsaveresult', data);
            },

            'processRightsChange': function(data) {
                $me.trigger('processrightschange', data);
            },

            'refreshHistory': function(data) {
                $me.trigger('refreshhistory', data);
            },

            'setHistoryData': function(data) {
                $me.trigger('sethistorydata', data);
            },

            'setEmailAddresses': function(data) {
                $me.trigger('setemailaddresses', data);
            },

            'setActionLink': function (data) {
                $me.trigger('setactionlink', data.url);
            },

            'processMailMerge': function(data) {
                $me.trigger('processmailmerge', data);
            },

            'downloadAs': function(data) {
                $me.trigger('downloadas', data);
            },

            'processMouse': function(data) {
                $me.trigger('processmouse', data);
            },

            'internalCommand': function(data) {
                $me.trigger('internalcommand', data);
            },

            'resetFocus': function(data) {
                $me.trigger('resetfocus', data);
            },

            'setUsers': function(data) {
                $me.trigger('setusers', data);
            },

            'showSharingSettings': function(data) {
                $me.trigger('showsharingsettings', data);
            },

            'setSharingSettings': function(data) {
                $me.trigger('setsharingsettings', data);
            },

            'insertImage': function(data) {
                $me.trigger('insertimage', data);
            },

            'setMailMergeRecipients': function(data) {
                $me.trigger('setmailmergerecipients', data);
            },

            'setRevisedFile': function(data) {
                $me.trigger('setrevisedfile', data);
            }
        };

        var _postMessage = function(msg) {
            // TODO: specify explicit origin
            if (window.parent && window.JSON) {
                msg.frameEditorId = window.frameEditorId;
                window.parent.postMessage(window.JSON.stringify(msg), "*");
            }
        };

        var _getPropByPath = function(obj, path) {
            var tempObj = obj;
            path = path.replace(/\[(\w+)\]/g, '.$1');
            path = path.replace(/^\./, '');

            var keyArr = path.split('.');
            var i = 0;

            for (var len = keyArr.length; i < len - 1; ++i) {
                var key = keyArr[i];
                if (key in tempObj) {
                    tempObj = tempObj[key];
                } else {
                    throw new Error('[iView warn]: please transfer a valid prop path to form item!');
                }
            }
            return tempObj[keyArr[i]];
        }

        var _isString = function(arg) {
            return Object.prototype.toString.call(arg) === '[object String]';
        }

        var _onMessage = function(msg) {
            // TODO: check message origin
            var data = msg.data;
            if (!_isString(data) || !window.JSON) {
                return;
            }

            var cmd, handler;

            try {
                cmd = window.JSON.parse(data)
            } catch(e) {
                cmd = '';
            }

            if (cmd) {
                if (cmd.source === 'bamboo') {
                    // cmd.data is --- eventObject = {
                    //     target: 'DE.controllers.xxx.xxx',
                    //     index:  0,
                    //     eventName: 'click',
                    //     eventParameters: [
                    //         {value: 'DE.xxx', description: 'menu'},
                    //         {value: 'DE.xxx.items', index: '#argumentsConfig.0', description: 'item' },
                    //         {value: true, description: 'state'}
                    //     ],
                    //     argumentsConfig: [ // sdk调用方需要传递的参数描述信息
                    //         {type: Boolean, default: true, items: [true, false], description: 'true:竖向,false:横向'}
                    //     ],
                    // };
                    if (cmd.data.type === 'method') {
                        var target = _getPropByPath(window, cmd.data.target);
                        if (!target || !target[cmd.data.methodName]) {
                            throw new Error(cmd.data.target + ' not found');
                        }
                        target[cmd.data.methodName].apply(target, cmd.data.parameters);
                        return;
                    }
                    // event
                    if (cmd.data.selectors) {
                        if (cmd.data.index) {
                            var target = document.querySelectorAll(cmd.data.selectors);
                            if (!target || target.length < cmd.data.index) {
                                // 不做操作
                                throw new Error(cmd.data.selectors + '['+cmd.data.index+'] not found');
                            }
                            target[cmd.data.index][cmd.data.eventName]();
                        } else {
                            var target = document.querySelector(cmd.data.selectors);
                            if (!target) {
                                // 不做操作
                                throw new Error(cmd.data.selectors + ' not found');
                            }
                            target[cmd.data.eventName]();
                        }
                        return;
                    }

                    var target = _getPropByPath(window, cmd.data.target);
                    if (!target) {
                        throw new Error(cmd.data.target + ' not found');
                    }
                    if (target.disabled) {
                        return;
                    }
                    target.doToggle && target.doToggle();
                    // if (target.options.hint) {
                    //     var tip = target.btnEl.data('bs.tooltip');
                    //     if (tip) {
                    //         if (tip.dontShow===undefined)
                    //             tip.dontShow = true;
                    //         tip.hide();
                    //     }
                    // }
                    cmd.data.target = target;
                    var parameters = [cmd.data.eventName];
                    if (cmd.data.eventParameters) {
                        for (var i = 0; i < cmd.data.eventParameters.length; i++) {
                            var eventParam = cmd.data.eventParameters[i];
                            if (_isString(eventParam.value) && eventParam.value.indexOf('#') === 0) {
                                var ref = eventParam.value.substring(1);
                                var arg = _getPropByPath(cmd.data, ref);
                                if (!arg) {
                                    throw new Error(eventParam.value + ' not found');
                                }
                                if (eventParam.index !== undefined) {
                                    arg = arg[eventParam.index];
                                }
                                parameters.push(arg);
                            } else {
                                parameters.push(eventParam.value);
                            }
                        }
                    } else if (cmd.data.index === undefined) {
                        parameters.push(target);
                    } else {
                        target = target[cmd.data.index];
                        parameters.push(target);
                    }
                    target.trigger.apply(target, parameters);
                } else {
                    //
                    handler = commandMap[cmd.command];
                    if (handler) {
                        handler.call(this, cmd.data);
                    }
                }

            }

        };

        var fn = function(e) { _onMessage(e); };

        if (window.attachEvent) {
            window.attachEvent('onmessage', fn);
        } else {
            window.addEventListener('message', fn, false);
        }

        return {

            appReady: function() {
                _postMessage({ event: 'onAppReady' });
            },

            requestEditRights: function() {
                _postMessage({ event: 'onRequestEditRights' });
            },

            requestHistory: function() {
                _postMessage({ event: 'onRequestHistory' });
            },

            requestHistoryData: function(revision) {
                _postMessage({
                    event: 'onRequestHistoryData',
                    data: revision
                });
            },

            requestRestore: function(version, url) {
                _postMessage({
                    event: 'onRequestRestore',
                    data: {
                        version: version,
                        url: url
                    }
                });
            },

            requestEmailAddresses: function() {
                _postMessage({ event: 'onRequestEmailAddresses' });
            },

            requestStartMailMerge: function() {
                _postMessage({event: 'onRequestStartMailMerge'});
            },

            requestHistoryClose: function(revision) {
                _postMessage({event: 'onRequestHistoryClose'});
            },

            reportError: function(code, description) {
                _postMessage({
                    event: 'onError',
                    data: {
                        errorCode: code,
                        errorDescription: description
                    }
                });
            },

            reportWarning: function(code, description) {
                _postMessage({
                    event: 'onWarning',
                    data: {
                        warningCode: code,
                        warningDescription: description
                    }
                });
            },

            sendInfo: function(info) {
                _postMessage({
                    event: 'onInfo',
                    data: info
                });
            },

            setDocumentModified: function(modified) {
                _postMessage({
                    event: 'onDocumentStateChange',
                    data: modified
                });
            },

            internalMessage: function(type, data) {
                _postMessage({
                    event: 'onInternalMessage',
                    data: {
                        type: type,
                        data: data
                    }
                });
            },

            updateVersion: function() {
                _postMessage({ event: 'onOutdatedVersion' });
            },

            downloadAs: function(url) {
                _postMessage({
                    event: 'onDownloadAs',
                    data: url
                });
            },

            requestSaveAs: function(url, title) {
                _postMessage({
                    event: 'onRequestSaveAs',
                    data: {
                        url: url,
                        title: title
                    }
                });
            },

            collaborativeChanges: function() {
                _postMessage({event: 'onCollaborativeChanges'});
            },

            requestRename: function(title) {
                _postMessage({event: 'onRequestRename', data: title});
            },

            metaChange: function(meta) {
                _postMessage({event: 'onMetaChange', data: meta});
            },

            documentReady: function() {
                _postMessage({ event: 'onDocumentReady' });
            },

            requestClose: function() {
                _postMessage({event: 'onRequestClose'});
            },

            requestMakeActionLink: function (config) {
                _postMessage({event:'onMakeActionLink', data: config});
            },

            requestUsers:  function () {
                _postMessage({event:'onRequestUsers'});
            },

            requestSendNotify:  function (emails) {
                _postMessage({event:'onRequestSendNotify', data: emails});
            },

            requestInsertImage:  function () {
                _postMessage({event:'onRequestInsertImage'});
            },

            requestMailMergeRecipients:  function () {
                _postMessage({event:'onRequestMailMergeRecipients'});
            },

            requestCompareFile:  function () {
                _postMessage({event:'onRequestCompareFile'});
            },

            requestSharingSettings:  function () {
                _postMessage({event:'onRequestSharingSettings'});
            },

            // add by xialiang@20200724#hyperlink click event
            requestHyperlinkClick: function (config) {
                _postMessage({event:'onHyperlinkClick', data: config});
            },

            on: function(event, handler){
                var localHandler = function(event, data){
                    handler.call(me, data)
                };

                $me.on(event, localHandler);
            }
        }

    })();

Common.Bamboo = new(function() {
    var me = this,
        $me = $(me);

    var commandMap = {
        'init': function(data) {
            $me.trigger('init', data);
        },

        'openDocument': function(data) {
            $me.trigger('opendocument', data);
        },

        'showMessage': function(data) {
            $me.trigger('showmessage', data);
        },

        'applyEditRights': function(data) {
            $me.trigger('applyeditrights', data);
        },

        'processSaveResult': function(data) {
            $me.trigger('processsaveresult', data);
        },

        'processRightsChange': function(data) {
            $me.trigger('processrightschange', data);
        },

        'refreshHistory': function(data) {
            $me.trigger('refreshhistory', data);
        },

        'setHistoryData': function(data) {
            $me.trigger('sethistorydata', data);
        },

        'setEmailAddresses': function(data) {
            $me.trigger('setemailaddresses', data);
        },

        'setActionLink': function (data) {
            $me.trigger('setactionlink', data.url);
        },

        'processMailMerge': function(data) {
            $me.trigger('processmailmerge', data);
        },

        'downloadAs': function(data) {
            $me.trigger('downloadas', data);
        },

        'processMouse': function(data) {
            $me.trigger('processmouse', data);
        },

        'internalCommand': function(data) {
            $me.trigger('internalcommand', data);
        },

        'resetFocus': function(data) {
            $me.trigger('resetfocus', data);
        },

        'setUsers': function(data) {
            $me.trigger('setusers', data);
        },

        'showSharingSettings': function(data) {
            $me.trigger('showsharingsettings', data);
        },

        'setSharingSettings': function(data) {
            $me.trigger('setsharingsettings', data);
        },

        'insertImage': function(data) {
            $me.trigger('insertimage', data);
        },

        'setMailMergeRecipients': function(data) {
            $me.trigger('setmailmergerecipients', data);
        },

        'setRevisedFile': function(data) {
            $me.trigger('setrevisedfile', data);
        }
    };

    var _postMessage = function(msg) {
        // TODO: specify explicit origin
        if (window.parent && window.JSON) {
            msg.frameEditorId = window.frameEditorId;
            window.parent.postMessage(window.JSON.stringify(msg), "*");
        }
    };

    var _onMessage = function(msg) {
        // TODO: check message origin
        var data = msg.data;
        if (Object.prototype.toString.apply(data) !== '[object String]' || !window.JSON) {
            return;
        }

        var cmd, handler;

        try {
            cmd = window.JSON.parse(data)
        } catch(e) {
            cmd = '';
        }

        if (cmd) {
            handler = commandMap[cmd.command];
            if (handler) {
                handler.call(this, cmd.data);
            }
            // bamboo

        }

    };

    var fn = function(e) { _onMessage(e); };

    if (window.attachEvent) {
        window.attachEvent('onmessage', fn);
    } else {
        window.addEventListener('message', fn, false);
    }

    return {

        appReady: function() {
            _postMessage({ event: 'onAppReady' });
        },

        requestEditRights: function() {
            _postMessage({ event: 'onRequestEditRights' });
        },

        requestHistory: function() {
            _postMessage({ event: 'onRequestHistory' });
        },

        requestHistoryData: function(revision) {
            _postMessage({
                event: 'onRequestHistoryData',
                data: revision
            });
        },

        requestRestore: function(version, url) {
            _postMessage({
                event: 'onRequestRestore',
                data: {
                    version: version,
                    url: url
                }
            });
        },

        requestEmailAddresses: function() {
            _postMessage({ event: 'onRequestEmailAddresses' });
        },

        requestStartMailMerge: function() {
            _postMessage({event: 'onRequestStartMailMerge'});
        },

        requestHistoryClose: function(revision) {
            _postMessage({event: 'onRequestHistoryClose'});
        },

        reportError: function(code, description) {
            _postMessage({
                event: 'onError',
                data: {
                    errorCode: code,
                    errorDescription: description
                }
            });
        },

        reportWarning: function(code, description) {
            _postMessage({
                event: 'onWarning',
                data: {
                    warningCode: code,
                    warningDescription: description
                }
            });
        },

        sendInfo: function(info) {
            _postMessage({
                event: 'onInfo',
                data: info
            });
        },

        setDocumentModified: function(modified) {
            _postMessage({
                event: 'onDocumentStateChange',
                data: modified
            });
        },

        internalMessage: function(type, data) {
            _postMessage({
                event: 'onInternalMessage',
                data: {
                    type: type,
                    data: data
                }
            });
        },

        updateVersion: function() {
            _postMessage({ event: 'onOutdatedVersion' });
        },

        downloadAs: function(url) {
            _postMessage({
                event: 'onDownloadAs',
                data: url
            });
        },

        requestSaveAs: function(url, title) {
            _postMessage({
                event: 'onRequestSaveAs',
                data: {
                    url: url,
                    title: title
                }
            });
        },

        collaborativeChanges: function() {
            _postMessage({event: 'onCollaborativeChanges'});
        },

        requestRename: function(title) {
            _postMessage({event: 'onRequestRename', data: title});
        },

        metaChange: function(meta) {
            _postMessage({event: 'onMetaChange', data: meta});
        },

        documentReady: function() {
            _postMessage({ event: 'onDocumentReady' });
        },

        requestClose: function() {
            _postMessage({event: 'onRequestClose'});
        },

        requestMakeActionLink: function (config) {
            _postMessage({event:'onMakeActionLink', data: config});
        },

        requestUsers:  function () {
            _postMessage({event:'onRequestUsers'});
        },

        requestSendNotify:  function (emails) {
            _postMessage({event:'onRequestSendNotify', data: emails});
        },

        requestInsertImage:  function () {
            _postMessage({event:'onRequestInsertImage'});
        },

        requestMailMergeRecipients:  function () {
            _postMessage({event:'onRequestMailMergeRecipients'});
        },

        requestCompareFile:  function () {
            _postMessage({event:'onRequestCompareFile'});
        },

        requestSharingSettings:  function () {
            _postMessage({event:'onRequestSharingSettings'});
        },

        on: function(event, handler){
            var localHandler = function(event, data){
                handler.call(me, data)
            };

            $me.on(event, localHandler);
        }
    }

})();
