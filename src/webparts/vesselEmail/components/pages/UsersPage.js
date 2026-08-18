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
exports.renderUsersPage = renderUsersPage;
var React = __importStar(require("react"));
var constants_1 = require("../constants");
function renderUsersPage(host) {
    var _a = host.state, usersList = _a.usersList, userSearch = _a.userSearch;
    var filtered = usersList.filter(function (u) {
        return u.name.toLowerCase().indexOf(userSearch.toLowerCase()) !== -1 ||
            u.email.toLowerCase().indexOf(userSearch.toLowerCase()) !== -1;
    });
    return (React.createElement("div", { style: { display: 'flex', flexDirection: 'column', gap: 16 } },
        React.createElement("div", { style: { display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 12 } },
            React.createElement("div", null,
                React.createElement("h2", { style: { margin: 0, fontSize: 20, fontWeight: 700, color: '#0f172a' } }, "User Management"),
                React.createElement("p", { style: { margin: '4px 0 0', fontSize: 13, color: '#64748b' } }, "Manage users and their roles.")),
            React.createElement("button", { onClick: function () { return alert('Add User Modal'); }, style: {
                    background: '#0078d4', color: '#fff', border: 'none', borderRadius: 6,
                    padding: '8px 18px', fontSize: 13, fontWeight: 600, cursor: 'pointer',
                    display: 'inline-flex', alignItems: 'center', gap: 6,
                } }, "\uFF0B Add User")),
        React.createElement("div", { style: { background: '#fff', borderRadius: 8, padding: 12, border: '1px solid #e2e8f0', display: 'flex', gap: 12 } },
            React.createElement("input", { type: "text", placeholder: "Search users", value: userSearch, onChange: function (e) { return host.setState({ userSearch: e.target.value }); }, style: { padding: '6px 12px', borderRadius: 6, border: '1px solid #cbd5e1', fontSize: 13, outline: 'none', width: 240 } })),
        React.createElement("div", { style: { background: '#fff', borderRadius: 8, border: '1px solid #e2e8f0', overflow: 'hidden' } },
            React.createElement("table", { style: { width: '100%', borderCollapse: 'collapse', fontSize: 13 } },
                React.createElement("thead", null,
                    React.createElement("tr", { style: { background: '#f8fafc', borderBottom: '1px solid #e2e8f0', color: '#64748b', fontSize: 11, textTransform: 'uppercase', textAlign: 'left' } },
                        React.createElement("th", { style: { padding: '10px 16px' } }, "Name"),
                        React.createElement("th", { style: { padding: '10px 16px' } }, "Email"),
                        React.createElement("th", { style: { padding: '10px 16px' } }, "Role"),
                        React.createElement("th", { style: { padding: '10px 16px' } }, "Status"),
                        React.createElement("th", { style: { padding: '10px 16px' } }, "Last Login"))),
                React.createElement("tbody", null, filtered.map(function (usr) { return (React.createElement("tr", { key: usr.id, style: { borderBottom: '1px solid #f1f5f9' } },
                    React.createElement("td", { style: { padding: '12px 16px', fontWeight: 600, color: '#1e293b', display: 'flex', alignItems: 'center', gap: 10 } },
                        React.createElement("div", { style: { width: 28, height: 28, borderRadius: '50%', background: '#e2e8f0', color: '#475569', fontSize: 12, fontWeight: 700, display: 'flex', alignItems: 'center', justifyContent: 'center' } }, usr.name.charAt(0)),
                        usr.name),
                    React.createElement("td", { style: { padding: '12px 16px', color: '#475569' } }, usr.email),
                    React.createElement("td", { style: { padding: '12px 16px', color: '#334155' } }, usr.role),
                    React.createElement("td", { style: { padding: '12px 16px' } }, (0, constants_1.badge)(usr.status === 'Active' ? 'green' : 'red', usr.status)),
                    React.createElement("td", { style: { padding: '12px 16px', color: '#64748b' } }, usr.lastLogin))); }))))));
}
