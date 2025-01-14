<template>
    <!-- <el-table :data="tableData" border style="width: 100%">
        <el-table-column :label="item" v-for="(item, index) in titleLength">
            <template #default="scope">
                {{ scope.row[index] }}
            </template>
</el-table-column>
</el-table> -->
    <el-upload class="upload-demo" drag action="#" :show-file-list="false" :http-request="handleUpload"
        :before-upload="handleChange">
        <el-icon class="el-icon--upload"><upload-filled /></el-icon>
        <div class="el-upload__text">
            点此或者拖拽文件 <em>这里</em>
        </div>
        <template #tip>
            <div class="el-upload__tip">
                只支持10MB以下的excel文件
            </div>
        </template>
    </el-upload>
</template>

<script setup>
import { ref, onMounted } from 'vue';
import { getExcel, uploadFile } from "@/api";
import { UploadFilled } from '@element-plus/icons-vue'

let tableData = ref([])
let titleLength = ref(0)

const getExcelData = async () => {
    let m = []
    const data = await getExcel()
    const d = data.data
    for (const key in d) {
        m.push(d[key])
    }
    tableData.value = m
    titleLength.value = m[0].length

}

const handleUpload = (rawFile) => {
    // ElMessage.error("上传图片最大不超过1MB!");
    console.log(rawFile);

    return true
}

const handleChange = async (file) => {
    const data = await uploadFile(file)
    console.log(data);

}

onMounted(async () => {
    // await getExcelData()
})
</script>



<style scoped lang="scss">
.common-layout {
    height: 100vh;

    .container {
        height: 100vh;

        .max-w {
            width: 100vw;
        }
    }
}
</style>