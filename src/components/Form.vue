<template>
    <form class="w-full max-w-xs m-auto">
        <div>
            <label for="name_input" class="block text-gray-700 text-sm font-bold mb-2">Name</label>
            <input
                v-model="reactiveForm.name"
                class="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline" 
                id="name_input" 
            >
        </div>
        <div>
            <label class="block">
            <span class="block text-gray-700 text-sm font-bold mb-2">Email address</span>
            <input type="email"
                v-model="reactiveForm.email"
                class="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline"
                placeholder="sample@example.com">
            </label>
        </div>
        <div>
            <label for="description" class="block text-gray-700 text-sm font-bold mb-2">Description</label>
            <textarea 
                v-model="reactiveForm.description"
                rows="3"
                class="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline" 
            ></textarea>
        </div>

        <button 
            type="button" 
            class="bg-blue-500 hover:bg-blue-700 text-white font-bold py-2 px-4 rounded focus:outline-none focus:shadow-outline" 
            @click="excelExport"
        >Excel Export</button>
    </form>
</template>

<script lang="ts">
import { reactive, defineComponent } from 'vue'
import exceljs from 'exceljs';

export default defineComponent({
  name: 'Form',

  setup: () => {
    const reactiveForm = reactive({
        name: undefined,
        email: undefined,
        description: undefined,
    })

    const excelExport = async (e: Event) => {
        e.preventDefault();

        const workbook = new exceljs.Workbook();
        const formUrl = './forms/form.xlsx';
        const form = await fetch(formUrl);
        const formBuffer = await form.arrayBuffer();

        await workbook.xlsx.load(formBuffer);
        const worksheet = workbook.getWorksheet('form1');
        worksheet.getCell('B3').value = reactiveForm.name;
        worksheet.getCell('C3').value = reactiveForm.email;
        worksheet.getCell('D3').value = reactiveForm.description;

        const writeBuffer = await workbook.xlsx.writeBuffer();

        const blob = new Blob([writeBuffer], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});

        const el = document.createElement('a');
        const urlApi = window.URL || window.webkitURL;
        const url = urlApi.createObjectURL(blob);
        el.href = url;
        el.download = 'form.xlsx';
        el.click();
        urlApi.revokeObjectURL(url);

    }
    return { reactiveForm, excelExport }
  }
})


</script>