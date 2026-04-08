import { SUPABASE_CONFIG } from './config.js';
const { createClient } = supabase;
const supabaseClient = createClient(SUPABASE_CONFIG.url, SUPABASE_CONFIG.anonKey);

export const DB = {
    async getTemplates(moduleName) {
        const { data, error } = await supabaseClient
            .from('doc_templates')
            .select('*')
            .eq('module', moduleName);
        if (error) throw error;
        return data || [];
    },

    async getTemplateDetails(templateId) {
        const [tRes, eRes] = await Promise.all([
            supabaseClient.from('doc_templates').select('*, doc_columns(*)').eq('id', templateId).single(),
            supabaseClient.from('doc_entries').select('*').eq('template_id', templateId).order('created_at', { ascending: false })
        ]);
        if (tRes.error) throw tRes.error;
        const template = tRes.data;
        template.doc_columns.sort((a, b) => a.display_order - b.display_order);
        return { template, entries: eRes.data || [] };
    },

    async saveEntry(templateId, content, id = null) {
        if (id) {
            return await supabaseClient.from('doc_entries').update({ content }).eq('id', id).select();
        }
        return await supabaseClient.from('doc_entries').insert([{ template_id: templateId, content }]).select();
    },

    async deleteEntry(id) {
        const { error } = await supabaseClient.from('doc_entries').delete().eq('id', id);
        if (error) throw error;
    },

    async createCategory(name, moduleName) {
        const { data, error } = await supabaseClient
            .from('doc_templates')
            .insert([{ name, module: moduleName }])
            .select();
        if (error) throw error;
        return data[0];
    },

    async deleteCategory(id) {
        const { error } = await supabaseClient.from('doc_templates').delete().eq('id', id);
        if (error) throw error;
    },

    async addColumn(templateId, name, type, order) {
        const { error } = await supabaseClient
            .from('doc_columns')
            .insert([{ template_id: templateId, column_name: name, column_type: type, display_order: order }]);
        if (error) throw error;
    }
};