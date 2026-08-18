"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
exports.renderSidebar = renderSidebar;
var React = __importStar(require("react"));
function renderSidebar(host) {
    var _a = host.state, view = _a.view, notificationsList = _a.notificationsList, sidebarCollapsed = _a.sidebarCollapsed, windowWidth = _a.windowWidth;
    var unreadNotifs = notificationsList.filter(function (n) { return !n.read; }).length;
    var toggleCollapsed = function () { return host.setState({ sidebarCollapsed: !sidebarCollapsed }); };
    var collapsed = sidebarCollapsed;
    var isNarrowScreen = (windowWidth || (typeof window !== 'undefined' ? window.innerWidth : 1200)) < 640;
    var expandedWidth = isNarrowScreen ? 190 : 230;
    var sidebarWidth = collapsed ? 52 : expandedWidth;
    // Collapsed: hamburger/menu icon. Expanded: chevron (‹) to collapse.
    var collapseToggleBtn = (React.createElement("button", { onClick: toggleCollapsed, title: collapsed ? 'Expand navigation' : 'Collapse navigation', style: {
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            width: 28, height: 28, borderRadius: 6, border: '1px solid #e5e7eb',
            background: '#fff', color: '#6b7280', cursor: 'pointer', flexShrink: 0,
            fontSize: 13, lineHeight: 1,
        } }, collapsed ? '☰' : String.fromCharCode(0xAB)));
    var navItems = [
        { id: 'dashboard', label: 'Home', icon: '🏠' },
        { id: 'list', label: 'Documents', icon: '📄' },
        { id: 'vessels', label: 'Vessels', icon: '🚢' },
        { id: 'templates', label: 'Templates', icon: '📑' },
        { id: 'approvals', label: 'Approvals', icon: '☑️' },
        { id: 'notifications', label: 'Notifications', icon: '🔔', badge: unreadNotifs > 0 ? unreadNotifs : undefined },
        { id: 'reports', label: 'Reports', icon: '📊' },
        { id: 'users', label: 'User Management', icon: '👥' },
        { id: 'settings', label: 'Settings', icon: '⚙️' },
    ];
    var auxLinks = [
        { id: 'bento_email', label: 'AI Bento Email', icon: '🤖', onClick: function () { return host._goToView('bento_email'); } },
        {
            id: 'bento_compose',
            label: 'Send Email',
            icon: '✉',
            onClick: function () {
                void host._goToView('bento_email');
                host.setState({
                    bentoComposeOpen: true,
                    bentoComposeMsg: null,
                    bentoComposeErr: null,
                    bentoComposeForm: {
                        vessel_name: '',
                        datasource_tag: 'mail',
                        subject_text: host._buildAutoSubject('', 'mail', ''),
                        body: '',
                        file: null,
                        existing_attachment: '',
                        recipient: '',
                    },
                });
            },
        },
        { id: 'recycle', label: 'Recycle bin', icon: '🗑️', onClick: function () { return host._goToView('recycle'); } },
        { id: 'archive', label: 'Archive', icon: '📦', onClick: function () { return host._goToView('archive'); } },
    ];
    return (React.createElement("div", { style: {
            width: sidebarWidth,
            minWidth: sidebarWidth,
            background: '#ffffff',
            borderRight: '1px solid #e0e0e0',
            display: 'flex',
            flexDirection: 'column',
            // Keep auxiliary modules directly below the main navigation instead of
            // pinning them to the bottom of the sidebar.
            justifyContent: 'flex-start',
            padding: '16px 0',
            boxSizing: 'border-box',
            userSelect: 'none',
            overflow: 'hidden',
            flexShrink: 0,
            transition: 'width 0.25s ease, min-width 0.25s ease',
        } },
        React.createElement("div", null,
            React.createElement("div", { style: {
                    display: 'flex',
                    alignItems: 'center',
                    gap: collapsed ? 0 : 12,
                    justifyContent: collapsed ? 'center' : 'flex-start',
                    padding: collapsed ? '0 0 20px' : '0 20px 20px',
                    borderBottom: '1px solid #f0f0f0',
                    transition: 'padding 0.25s ease, gap 0.25s ease',
                } },
                !collapsed && (React.createElement(React.Fragment, null,
                    React.createElement("div", { style: {
                            width: 36, height: 36, borderRadius: 8, background: '#0078d4',
                            color: '#fff', fontWeight: 700, fontSize: 16, display: 'flex',
                            alignItems: 'center', justifyContent: 'center', boxShadow: '0 2px 6px rgba(0,120,212,0.3)',
                            flexShrink: 0,
                        } }, "VD"),
                    React.createElement("div", { style: { flex: 1, minWidth: 0 } },
                        React.createElement("div", { style: { fontWeight: 700, fontSize: 14, color: '#111827', lineHeight: 1.2, whiteSpace: 'nowrap' } }, "Vessel Documents"),
                        React.createElement("div", { style: { fontSize: 12, color: '#6b7280', fontWeight: 500, whiteSpace: 'nowrap' } }, "Management")))),
                collapseToggleBtn),
            React.createElement("div", { style: { padding: collapsed ? '12px 8px' : '12px 10px', transition: 'padding 0.25s ease' } }, navItems.map(function (item) {
                var active = view === item.id;
                return (React.createElement("button", { key: item.id, onClick: function () { return host._goToView(item.id); }, title: collapsed ? item.label : undefined, style: {
                        position: 'relative',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: collapsed ? 'center' : 'space-between',
                        width: '100%',
                        padding: collapsed ? '9px 0' : '9px 14px',
                        marginBottom: 3,
                        borderRadius: 6,
                        border: 'none',
                        background: active ? '#eff6ff' : 'transparent',
                        color: active ? '#0078d4' : '#4b5563',
                        fontWeight: active ? 600 : 500,
                        fontSize: 13,
                        cursor: 'pointer',
                        textAlign: 'left',
                        transition: 'background 0.15s ease, color 0.15s ease, padding 0.25s ease',
                    } },
                    React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: collapsed ? 0 : 10, minWidth: 0 } },
                        React.createElement("span", { style: { fontSize: 16, flexShrink: 0 } }, item.icon),
                        React.createElement("span", { style: {
                                maxWidth: collapsed ? 0 : 150,
                                opacity: collapsed ? 0 : 1,
                                overflow: 'hidden',
                                whiteSpace: 'nowrap',
                                display: 'inline-block',
                                transition: 'max-width 0.25s ease, opacity 0.15s ease',
                            } }, item.label)),
                    item.badge !== undefined && (React.createElement("span", { style: {
                            position: collapsed ? 'absolute' : 'static',
                            top: collapsed ? -2 : undefined,
                            right: collapsed ? -2 : undefined,
                            background: '#ef4444', color: '#fff',
                            borderRadius: collapsed ? 8 : 10,
                            padding: collapsed ? '0 4px' : '1px 7px',
                            fontSize: collapsed ? 9 : 10,
                            fontWeight: 700,
                            lineHeight: collapsed ? '14px' : undefined,
                            transition: 'all 0.2s ease',
                        } }, item.badge))));
            }))),
        React.createElement("div", { style: {
                padding: collapsed ? '0 8px 12px' : '0 10px 12px',
                transition: 'padding 0.25s ease',
            } }, auxLinks.map(function (link) {
            var active = view === link.id;
            return (React.createElement("button", { key: link.id, onClick: link.onClick, title: collapsed ? link.label : undefined, style: {
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: collapsed ? 'center' : 'flex-start',
                    gap: collapsed ? 0 : 10,
                    width: '100%',
                    padding: collapsed ? '9px 0' : '8px 14px',
                    marginBottom: 2,
                    borderRadius: 6,
                    border: 'none',
                    background: active ? '#eff6ff' : 'transparent',
                    color: active ? '#0078d4' : '#6b7280',
                    fontSize: 12,
                    fontWeight: 500,
                    cursor: 'pointer',
                    whiteSpace: 'nowrap',
                    transition: 'background 0.15s ease, color 0.15s ease, padding 0.25s ease',
                } },
                React.createElement("span", { style: { fontSize: 15, flexShrink: 0 } }, link.icon),
                React.createElement("span", { style: {
                        maxWidth: collapsed ? 0 : 150,
                        opacity: collapsed ? 0 : 1,
                        overflow: 'hidden',
                        whiteSpace: 'nowrap',
                        display: 'inline-block',
                        transition: 'max-width 0.25s ease, opacity 0.15s ease',
                    } }, link.label)));
        }))));
}
