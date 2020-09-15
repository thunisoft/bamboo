/**
 * BambooOffice word api
 *
 * @author   BambooOffice Team
 * @version  1.0.0-beta.0
 * @date     2020-07-22
 */
;(function(window, document) {

    var frameName = 'frameEditor';
    var bookmarkManager = {
        type: 'method',
        target: 'DE.controllers.Viewport.api',
        methodName: 'asc_GetBookmarksManager',
    };
    var logicDocument = {
        type: 'method',
        target: 'DE.controllers.Viewport.api',
        methodName: 'asc_GetLogicDocument'
    };
    // 菜单相关操作
    var map = { // default is target@eventName
        // ============================ header ============================
        save: 'DE.controllers.Toolbar.toolbar.btnSave@click', // 保存
        print: 'DE.controllers.Toolbar.toolbar.btnPrint@click', // 打印
        undo: 'DE.controllers.Toolbar.toolbar.btnUndo@click', // 撤销
        redo: 'DE.controllers.Toolbar.toolbar.btnRedo@click', // 重做
        // fitPage: 'DE.controllers.Viewport.header.mnuitemFitPage@click', // 适合页面
        // fitWidth: 'DE.controllers.Viewport.header.mnuitemFitWidth@click', // 适合宽度

        // ============================ 主页 ============================
        copy: 'DE.controllers.Toolbar.toolbar.btnCopy@click', // 复制
        paste: 'DE.controllers.Toolbar.toolbar.btnPaste@click', // 粘贴

        incFontSize: 'DE.controllers.Toolbar.toolbar.btnIncFontSize@click', // 增大字号
        decFontSize: 'DE.controllers.Toolbar.toolbar.btnDecFontSize@click', // 减小字号
        bold: 'DE.controllers.Toolbar.toolbar.btnBold@click', // 粗体
        italic: 'DE.controllers.Toolbar.toolbar.btnItalic@click', // 倾斜
        underline: 'DE.controllers.Toolbar.toolbar.btnUnderline@click', // 下划线
        strikeout: 'DE.controllers.Toolbar.toolbar.btnStrikeout@click', // 删除线
        btnSuperscript: 'DE.controllers.Toolbar.toolbar.btnSuperscript@click', // 上标
        btnSubscript: 'DE.controllers.Toolbar.toolbar.btnSubscript@click',  // 下标

        incLeftOffset: 'DE.controllers.Toolbar.toolbar.btnIncLeftOffset@click', // 增加缩进
        decLeftOffset: 'DE.controllers.Toolbar.toolbar.btnDecLeftOffset@click', // 增加缩进
        lineSpace: {
            type: 'event',
            selectors: '#id-toolbar-btn-linespace [type=menuitem]',
            index: '#argumentsConfig.0',
            eventName: 'click',
            argumentsConfig: [ // sdk调用方需要传递的参数描述信息
                {type: Number, default: 1.15, items: [1.0, 1.15, 1.5, 2.0, 2.5, 3.0], description: '段线间距'}
            ]
        },
        alignLeft: 'DE.controllers.Toolbar.toolbar.btnAlignLeft@click', // 左对齐
        alignCenter: 'DE.controllers.Toolbar.toolbar.btnAlignCenter@click', // 居中对齐
        alignRight: 'DE.controllers.Toolbar.toolbar.btnAlignRight@click', // 右对齐
        alignJust: 'DE.controllers.Toolbar.toolbar.btnAlignJust@click', // 正当, 两端对齐

        clearStyle: 'DE.controllers.Toolbar.toolbar.btnClearStyle@click', // 清除样式

        // ============================ 插入 ============================
        insertBlankPage: 'DE.controllers.Toolbar.toolbar.btnBlankPage@click',
        // insertTable: '', // TODO

        // ============================ 布局 ============================
        // pageMargins: 'DE.controllers.Toolbar.toolbar.btnPageMargins.menu', //
        pageOrient: { // 布局: 纵向, 横向
            type: 'event',
            selectors: '#tlbtn-pageorient [type=menuitem]',
            index: '#argumentsConfig.0',
            eventName: 'click',
            argumentsConfig: [ // sdk调用方需要传递的参数描述信息
                {type: Boolean, default: true, items: [true, false], description: 'true:纵向,false:横向'}
            ]
        },
        // btnPageSize: 'DE.controllers.Toolbar.toolbar.btnPageMargins.menu', //

        // ============================ 参考 ============================
        // ============================ 协作 ============================

        // %%%%%%%%%%%%%%%%%%%%%%%%%%%% 函数 %%%%%%%%%%%%%%%%%%%%%%%%%%%%
        goToBookmark: {
            type: 'method',
            target: bookmarkManager,
            methodName: 'asc_GoToBookmark',
            // parameters: [],
            argumentsConfig: [ // sdk调用方需要传递的参数描述信息
                {type: String, required: true, description: '书签名称'}
            ]
        },
        selectBookmark: {
            type: 'method',
            target: bookmarkManager,
            methodName: 'asc_SelectBookmark',
            argumentsConfig: [
                {type: String, required: true, description: '书签名称'}
            ]
        },
        getAllBookmarks: {
            type: 'method',
            target: bookmarkManager,
            methodName: 'asc_GetAllBookmarks'
        },
        getBookmarkByName: {
            type: 'method',
            target: bookmarkManager,
            methodName: 'asc_GetBookmarkByName2',
            argumentsConfig: [
                {type: String, required: true, description: '书签名称'}
            ]
        },
        addBookmark: {
            type: 'method',
            target: bookmarkManager,
            methodName: 'asc_AddBookmark',
            // parameters: [],
            argumentsConfig: [ // sdk调用方需要传递的参数描述信息
                {type: String, required: true, description: '书签名称'}
            ]
        },
        removeBookmark: {
            type: 'method',
            target: bookmarkManager,
            methodName: 'asc_RemoveBookmark',
            // parameters: [],
            argumentsConfig: [ // sdk调用方需要传递的参数描述信息
                {type: String, required: true, description: '书签名称'}
            ]
        },
        addText: {
            type: 'method',
            target: 'DE.controllers.Viewport.api',
            methodName: 'pluginMethod_InputText',
            // parameters: [],
            argumentsConfig: [ // sdk调用方需要传递的参数描述信息
                {type: String, required: true, description: '文本内容'},
            ]
        },
        addImageUrl: {
            type: 'method',
            target: 'DE.controllers.Viewport.api',
            methodName: 'AddImageUrl',
            argumentsConfig: [
                {type: String, required: true, description: '图片地址'},
                // 图片配置，当前支持：
                // WrappingStyle：0-6控制图片的布局样式
                // x：图片相对于当前页的横坐标
                // y：图片相对于当前页的纵坐标
                {type: Object, required: true, description: '图片配置'}
            ]
        },
        getCurrentPageNumber: {
            type: 'method',
            target: 'DE.controllers.Viewport.api',
            methodName: 'getCurrentPage'
        },
        getPageCount: {
            type: 'method',
            target: 'DE.controllers.Viewport.api',
            methodName: 'getCountPages'
        },
        getCursorPosition: {
            type: 'method',
            target: logicDocument,
            methodName: 'GetCursorPosXY'
        },
        gotoPage: {
            type: 'method',
            target: 'DE.controllers.Viewport.api',
            methodName: 'goToPage',
            argumentsConfig: [
                {type: Number, required: true, description: '页码，0-base'}
            ]
        },
        gotoHeader: {
            type: 'method',
            target: 'DE.controllers.Viewport.api',
            methodName: 'GoToHeader',
            argumentsConfig: [
                {type: Number, required: true, description: '页码，0-base'}
            ]
        },
        gotoFooter: {
            type: 'method',
            target: 'DE.controllers.Viewport.api',
            methodName: 'GoToFooter',
            argumentsConfig: [
                {type: Number, required: true, description: '页码，0-base'}
            ]
        },
        exitHeaderFooter: {
            type: 'method',
            target: 'DE.controllers.Viewport.api',
            methodName: 'ExitHeader_Footer',
            argumentsConfig: [
                {type: Number, required: true, description: '页码，0-base'}
            ]
        },
        moveCursorToEndOfLine: {
            type: 'method',
            target: logicDocument,
            methodName: 'MoveCursorToEndOfLine'
        },
        moveCursorToStartOfLine: {
            type: 'method',
            target: logicDocument,
            methodName: 'MoveCursorToStartOfLine'
        },
        moveCursorUp: {
            type: 'method',
            target: logicDocument,
            methodName: 'MoveCursorUp'
        },
        moveCursorDown: {
            type: 'method',
            target: logicDocument,
            methodName: 'MoveCursorDown'
        },
        moveCursorToStartPos: {
            type: 'method',
            target: logicDocument,
            methodName: 'MoveCursorToStartPos'
        },
        moveCursorToEndPos: {
            type: 'method',
            target: logicDocument,
            methodName: 'MoveCursorToEndPos'
        },
        moveCursorToPageStart: {
            type: 'method',
            target: logicDocument,
            methodName: 'MoveCursorToPageStart'
        },
        moveCursorToPageEnd: {
            type: 'method',
            target: logicDocument,
            methodName: 'MoveCursorToPageEnd'
        },
        moveCursorLeft: {
            type: 'method',
            target: logicDocument,
            methodName: 'MoveCursorLeft'
        },
        moveCursorRight: {
            type: 'method',
            target: logicDocument,
            methodName: 'MoveCursorRight'
        },
        selectAll: {
            type: 'method',
            target: logicDocument,
            methodName: 'SelectAll'
        },
        getSelectedText: {
            type: 'method',
            target: logicDocument,
            methodName: 'GetSelectedText'
        },
        findTextObject: {
            type: 'method',
            target: 'DE.controllers.Viewport.api',
            methodName: 'findTextObject',
            argumentsConfig: [
                {type: String, required: true, description: '文本内容'},
                {type: Boolean, required: true, description: '是否从第一个文本进行搜索'}
            ]
        },
        paste: {
            type: 'method',
            target: logicDocument,
            methodName: 'paste',
            argumentsConfig: [
                {type: Boolean, default: false, description: '是否带格式进行粘贴'}
            ]
        }
    }

    var _getBasePath = function() {
        var scripts = document.getElementsByTagName('script'),
            match;
        for (var i = scripts.length - 1; i >= 0; i--) {
            match = scripts[i].src.match(/(.*)api\/documents\/bamboo.js/i);
            if (match) {
                return match[1];
            }
        }

        return '';
    }

    var _importDocsAPI = function () {
        if (window.DocsAPI) {
            return;
        }
        var script = document.createElement('script');
        script.type = 'text/javascript';
        script.src = _getBasePath() + 'api/documents/api.js';
        document.body.appendChild(script);
    }

    var _awaitDocsAPI = function (onload) {
        if (window.DocsAPI) {
            onload()
        } else {
            setTimeout(function () {
                _awaitDocsAPI(onload)
            }, 500)
        }
    }

    var _isString = function(arg) {
        return Object.prototype.toString.call(arg) === '[object String]';
    }
    var _isFunction = function(arg) {
        return Object.prototype.toString.call(arg) === '[object Function]';
    }
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
                throw new Error('please transfer a valid prop path to form item!');
            }
        }
        return tempObj[keyArr[i]];
    }
    var _deepCopy = function (source) {
        var t = Object.prototype.toString.call(source)
        var o
        if (t === '[object Array]') {
            o = []
        } else if (t === '[object Object]') {
            o = {}
        } else {
            return source
        }
        if (t === '[object Array]') {
            for (var i = 0; i < source.length; i++) {
                o.push(_deepCopy(source[i]))
            }
        } else if (t === '[object Object]') {
            for (var key in source) {
                o[key] = _deepCopy(source[key])
            }
        }
        return o
    }

    var _translateIndex = function(eventObject, index, callArguments) {
        if (index === undefined || !_isString(index)) {
            return
        }
        if (index.indexOf('#argumentsConfig.') === 0) {
            var argConfigIndex = ~~index.substring('#argumentsConfig.'.length);
            var argConfig = eventObject.argumentsConfig[argConfigIndex];
            var callArg = callArguments[argConfigIndex] === undefined ? argConfig.default : argConfig.type(callArguments[argConfigIndex]);
            var itemIndex = argConfig.items.indexOf(callArg);
            if (itemIndex === -1) {
                throw new Error('Illegal argument, index: ' + argConfigIndex + ', value: ' + callArguments[argConfigIndex] + '. correct is ' + JSON.stringify(argConfig.items));
            }
            return itemIndex;
        }
    }
    var _callbackMap = {};
    var _callbackIndex = 0;
    var _apiInit = function (placeholderId, config) {
        config.events = config.events || {}
        var self = this;
        var onAppReady = config.events.onAppReady;
        config.events.onAppReady = function () {
            self.frame = document.getElementsByName(frameName)[0];
            self.frameWindow = self.frame.contentWindow;
            try {
                self.frameDocument = self.frame.contentDocument;
            } catch(e) {}
            onAppReady && onAppReady();
        }
        config.events.onCallbackEntrypoint = function(event) {
            var data = event.data;
            var callback = _callbackMap[data.callbackName];
            if (callback) {
                callback(data.result);
                delete _callbackMap[data.callbackName];
            }
        }
        _awaitDocsAPI(function() {
            var docEditor = new window.DocsAPI.DocEditor(placeholderId, config);
            for (var key in docEditor) {
                BambooAPI.prototype[key] = docEditor[key];
            }
        })
    }

    _importDocsAPI();

    function BambooAPI() {}

    BambooAPI.create = function(placeholderId, config) {
        var apiBasePath = _getBasePath() // api所在path
        var windowBasePath = window.location.port ? (window.location.hostname + ':' + window.location.port + '/') : window.location.hostname + '/';
        var sameOrigin = apiBasePath.substring(apiBasePath.indexOf('://')+3).indexOf(windowBasePath) === 0
        return sameOrigin ? new SameOriginBambooAPI(placeholderId, config) : new CrossOriginBambooAPI(placeholderId, config);
    }

    for (var key in map) {
        (function(fnName) {
            BambooAPI.prototype[fnName] = function () {
                var callObject = map[fnName];
                if (_isString(callObject)) {
                    if (callObject.indexOf('@') > 0) { // event
                        callObject = {
                            type: 'event',
                            target: callObject.substring(0, callObject.indexOf('@')),
                            eventName: callObject.substring(callObject.indexOf('@') + 1),
                            // parameters: [
                            //   {value: '#target', description: 'menu'},
                            //   {value: '#target.items', index: '#argumentsConfig.0', description: 'item' },
                            //   {value: true, description: 'state'}
                            // ],
                            // argumentsConfig: [ // sdk调用方需要传递的参数描述信息
                            //   {type: Boolean, default: true, items: [true, false], description: 'true:纵向,false:横向'}
                            // ]
                        };
                    } else {
                        callObject = {
                            type: 'method',
                            target: callObject,
                            // argumentsConfig: [ // sdk调用方需要传递的参数描述信息
                            //   {type: Boolean, default: true, required: false, description: 'true:纵向,false:横向'}
                            // ]
                        };
                    }
                } else {
                    callObject = _deepCopy(callObject);
                }

                if (callObject.type === 'event') {
                    var targetIndex = _translateIndex(callObject, callObject.index, arguments);
                    if (targetIndex !== undefined) {
                        callObject.index = targetIndex;
                    }
                    if (callObject.parameters) {
                        for (var i = 0; i < callObject.parameters.length; i++) {
                            var eventParam = callObject.parameters[i];
                            var itemIndex = _translateIndex(callObject, eventParam.index, arguments);
                            if (itemIndex !== undefined) {
                                eventParam.index = itemIndex;
                            }
                        }
                    }
                    this.trigger(callObject);
                } else { // method
                    callObject.parameters = callObject.parameters || [];
                    callObject.argumentsConfig = callObject.argumentsConfig || [];
                    var callback;
                    for (var i = 0; i < callObject.argumentsConfig.length; i++) {
                        var argConfig = callObject.argumentsConfig[i];
                        var callArg = arguments[i];
                        var isFn = _isFunction(callArg);
                        var isUndefined = callArg === undefined;
                        if (isFn || isUndefined) {
                            if (argConfig.required) {
                                throw new Error('Argument ' + i + ' is required, correct is ' + JSON.stringify(argConfig));
                            }
                            if (isUndefined) {
                                callArg = argConfig.default;
                            } else if (isFn) {
                                // 竹简API不会支持函数类的参数, 所以遇到函数就作为结果的回调函数
                                callback = callArg;
                                break;
                            }
                        } else {
                            callArg = argConfig.type(callArg);
                        }
                        callObject.parameters.push(callArg);
                    }
                    // 如果额外传递了一个参数, 并且是函数, 则作为callback
                    callback = callback || arguments[callObject.argumentsConfig.length];
                    if (_isFunction(callback)) {
                        var callbackName = fnName + '_callback_' + (_callbackIndex++);
                        _callbackMap[callbackName] = callback;
                        callObject.callbackName = callbackName;
                    }
                    return this.invoke(callObject);
                }
            }
        })(key);
    }

    window.BambooAPI = BambooAPI;
    /**
     * 同源对象
     * @param frame iframe对象
     * @constructor
     */
    function SameOriginBambooAPI (placeHolderId, config) {
        _apiInit.call(this, placeHolderId, config);
    }

    SameOriginBambooAPI.prototype = new BambooAPI();
    SameOriginBambooAPI.constructor = SameOriginBambooAPI;

    SameOriginBambooAPI.prototype.invoke = function(callObject) {
        var target;
        if (!_isString(callObject.target)) {
            target = this.invoke(callObject.target);
        } else {
            target = _getPropByPath(this.frameWindow, callObject.target);
        }
        if (!target || !target[callObject.methodName]) {
            throw new Error(callObject.target + ' not found');
        }
        var result = target[callObject.methodName].apply(target, callObject.parameters);
        if (callObject.callbackName) {
            var callback = _callbackMap[callObject.callbackName];
            if (callback) {
                callback(result);
                delete _callbackMap[callObject.callbackName];
                return;
            }
        }
        return result;
    }

    SameOriginBambooAPI.prototype.trigger = function(callObject) {
        if (callObject.selectors) {
            if (callObject.index) {
                var target = this.frameDocument.querySelectorAll(callObject.selectors);
                if (!target || target.length < callObject.index) {
                    // 不做操作
                    throw new Error(callObject.selectors + '['+callObject.index+'] not found');
                }
                target[callObject.index][callObject.eventName]();
            } else {
                var target = this.frameDocument.querySelector(callObject.selectors);
                if (!target) {
                    // 不做操作
                    throw new Error(callObject.selectors + ' not found');
                }
                target[callObject.eventName]();
            }
            return;
        }
        //
        var target = _getPropByPath(this.frameWindow, callObject.target);
        if (!target) {
            throw new Error(callObject.target + ' not found');
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
        callObject.target = target;
        var parameters = [callObject.eventName];
        if (callObject.parameters) {
            for (var i = 0; i < callObject.parameters.length; i++) {
                var eventParam = callObject.parameters[i];
                if (_isString(eventParam.value) && eventParam.value.indexOf('#') === 0) {
                    var ref = eventParam.value.substring(1);
                    var arg = _getPropByPath(callObject, ref);
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
        } else if (callObject.index === undefined) {
            parameters.push(target);
        } else {
            target = target[callObject.index];
            parameters.push(target);
        }
        target.trigger.apply(target, parameters);
    }
    /**
     * 跨域对象
     * @param frame iframe对象
     * @constructor
     */
    function CrossOriginBambooAPI (placeHolderId, config) {
        _apiInit.call(this, placeHolderId, config)
    }

    CrossOriginBambooAPI.prototype = new BambooAPI();
    CrossOriginBambooAPI.constructor = CrossOriginBambooAPI;

    CrossOriginBambooAPI.prototype.invoke = function(callObject) {
        var message = {
            source: 'bamboo',
            data: callObject
        };
        // 为了兼容旧版IE, 先序列化为字符串吧
        this.frameWindow.postMessage(JSON.stringify(message), '*');
    }
    CrossOriginBambooAPI.prototype.trigger = CrossOriginBambooAPI.prototype.invoke;

})(window, document);
