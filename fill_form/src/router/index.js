import { createRouter, createWebHashHistory } from 'vue-router';
const router = createRouter({ history: createWebHashHistory(), routes: [
  { path: '/', name: 'home', component: () => import('../views/HomeView.vue') },
  { path: '/knowledge', name: 'knowledge', component: () => import('../views/KnowledgeView.vue') },
  { path: '/forms/:formId/review', name: 'form-review', component: () => import('../views/FormReviewView.vue') },
  { path: '/history', name: 'history', component: () => import('../views/HistoryView.vue') },
  { path: '/settings', name: 'settings', component: () => import('../views/SettingsView.vue') },
  { path: '/writing-style', redirect: '/settings' },
] });
export default router;
