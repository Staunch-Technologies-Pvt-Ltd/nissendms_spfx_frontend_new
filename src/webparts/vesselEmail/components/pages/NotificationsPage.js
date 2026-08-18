"use strict";
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
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
exports.renderNotificationsPage = renderNotificationsPage;
var React = __importStar(require("react"));
var constants_1 = require("../constants");
function renderNotificationsPage(host) {
    var _a = host.state, notificationsList = _a.notificationsList, notificationFilter = _a.notificationFilter;
    var filtered = notificationsList.filter(function (n) { return notificationFilter === 'all' || !n.read; });
    return (React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 16 } },
        React.createElement("div", { style: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 12 } },
            React.createElement("div", null,
                React.createElement("h2", { style: { margin: 0, fontSize: 20, fontWeight: 700, color: '#0f172a' } }, "Notifications"),
                React.createElement("p", { style: { margin: '4px 0 0', fontSize: 13, color: '#64748b' } }, "Stay updated with important alerts and reminders.")),
            React.createElement("button", { onClick: function () {
                    var updated = notificationsList.map(function (n) { return (__assign(__assign({}, n), { read: true })); });
                    host.setState({ notificationsList: updated });
                }, style: { background: 'none', border: 'none', color: '#2563eb', fontSize: 13, fontWeight: 600, cursor: 'pointer' } }, "Mark all as read")),
        React.createElement("div", { style: { display: 'flex', gap: 8 } },
            React.createElement("button", { onClick: function () { return host.setState({ notificationFilter: 'all' }); }, style: {
                    border: 'none', background: notificationFilter === 'all' ? '#0078d4' : '#f1f5f9',
                    color: notificationFilter === 'all' ? '#fff' : '#475569', borderRadius: 16,
                    padding: '4px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer',
                } },
                "All (",
                notificationsList.length,
                ")"),
            React.createElement("button", { onClick: function () { return host.setState({ notificationFilter: 'unread' }); }, style: {
                    border: 'none', background: notificationFilter === 'unread' ? '#0078d4' : '#f1f5f9',
                    color: notificationFilter === 'unread' ? '#fff' : '#475569', borderRadius: 16,
                    padding: '4px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer',
                } },
                "Unread (",
                notificationsList.filter(function (n) { return !n.read; }).length,
                ")")),
        React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 10 } }, filtered.map(function (notif) { return (React.createElement("div", { key: notif.id, style: {
                background: notif.read ? '#fff' : '#f0f9ff',
                borderRadius: 8, border: '1px solid #e2e8f0', padding: '14px 18px',
                display: 'flex', justifyContent: 'space-between', alignItems: 'center',
            } },
            React.createElement("div", { style: { display: 'flex', alignItems: 'center', gap: 12 } },
                React.createElement("div", { style: {
                        width: 32, height: 32, borderRadius: '50%',
                        background: notif.type === 'alert' ? '#fde7e9' : notif.type === 'warning' ? '#fff4ce' : notif.type === 'success' ? '#dff6dd' : '#e1efff',
                        color: notif.type === 'alert' ? '#a4262c' : notif.type === 'warning' ? '#8a5700' : notif.type === 'success' ? '#107c10' : '#0078d4',
                        display: 'flex', alignItems: 'center', justifyContent: 'center', fontWeight: 700,
                    } }, notif.type === 'alert' ? '⚠' : notif.type === 'warning' ? '🔔' : notif.type === 'success' ? '✓' : 'ℹ'),
                React.createElement("div", null,
                    React.createElement("div", { style: { fontSize: 13, color: '#1e293b', fontWeight: notif.read ? 500 : 700 } }, notif.message),
                    React.createElement("div", { style: { fontSize: 11, color: '#64748b', marginTop: 2 } }, notif.timestamp))),
            React.createElement("div", null, (0, constants_1.badge)(notif.priority === 'High' ? 'red' : notif.priority === 'Medium' ? 'orange' : 'blue', notif.priority)))); }))));
}
