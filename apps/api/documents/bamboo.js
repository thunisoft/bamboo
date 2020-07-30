/**
 * BambooOffice word api
 *
 * @author   BambooOffice Team
 * @version  1.0.0-beta.0
 * @date     2020-07-22
 */
;(function(window, document) {

    var frameName = 'frameEditor';
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
        pageOrient: { // 布局: 竖向, 横向
            type: 'event',
            selectors: '#tlbtn-pageorient [type=menuitem]',
            index: '#argumentsConfig.0',
            eventName: 'click',
            argumentsConfig: [ // sdk调用方需要传递的参数描述信息
                {type: Boolean, default: true, items: [true, false], description: 'true:竖向,false:横向'}
            ]
        },
        // btnPageSize: 'DE.controllers.Toolbar.toolbar.btnPageMargins.menu', //

        // ============================ 参考 ============================
        // ============================ 协作 ============================

        // %%%%%%%%%%%%%%%%%%%%%%%%%%%% 函数 %%%%%%%%%%%%%%%%%%%%%%%%%%%%
        goToBookmark: {
            type: 'method',
            target: 'DE.controllers.Viewport.api.WordControl.m_oLogicDocument.BookmarksManager',
            methodName: 'GoToBookmark',
            // parameters: [],
            argumentsConfig: [ // sdk调用方需要传递的参数描述信息
                {type: String, required: true, description: '书签名称'}
            ]
        },

        addText: {
            type: 'method',
            target: 'DE.controllers.Viewport.api',
            methodName: 'AddText',
            // parameters: [],
            argumentsConfig: [ // sdk调用方需要传递的参数描述信息
                {type: String, required: true, description: '文本内容'}
            ]
        },
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
        var basePath = _getBasePath() // api所在path
        var sameOrigin = basePath.substring(basePath.indexOf('://')+3).indexOf(window.location.hostname + ':' + window.location.port) === 0
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
                            //   {type: Boolean, default: true, items: [true, false], description: 'true:竖向,false:横向'}
                            // ]
                        };
                    } else {
                        callObject = {
                            type: 'method',
                            target: callObject,
                            // argumentsConfig: [ // sdk调用方需要传递的参数描述信息
                            //   {type: Boolean, default: true, required: false, description: 'true:竖向,false:横向'}
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
                    if (callObject.argumentsConfig) {
                        for (var i = 0; i < callObject.argumentsConfig.length; i++) {
                            var argConfig = callObject.argumentsConfig[i];
                            var callArg = arguments[i];
                            if (argConfig.required && callArg === undefined) {
                                throw new Error('Argument ' + i + 'is required, correct is ' + JSON.stringify(argConfig));
                            }
                            callArg = callArg === undefined ? argConfig.default : argConfig.type(callArg);
                            callObject.parameters.push(callArg);
                        }
                    }
                    this.invoke(callObject);
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
        var target = _getPropByPath(this.frameWindow, callObject.target);
        if (!target || !target[callObject.methodName]) {
            throw new Error(callObject.target + ' not found');
        }
        target[callObject.methodName].apply(target, callObject.parameters);
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

    CrossOriginBambooAPI.prototype.trigger = function(callObject) {
        var message = {
            source: 'bamboo',
            data: callObject
        };
        // 为了兼容旧版IE, 先序列化为字符串吧
        this.frameWindow.postMessage(JSON.stringify(message), '*');
    }

    CrossOriginBambooAPI.prototype.invoke = CrossOriginBambooAPI.prototype.trigger
})(window, document);
