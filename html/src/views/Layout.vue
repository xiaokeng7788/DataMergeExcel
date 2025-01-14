<template>
    <div class="common-layout">
        <el-container class="container">
            <el-header class="max-w">Header</el-header>
            <el-main class="max-w">
                <Main></Main>
            </el-main>
            <el-footer class="max-w">Footer</el-footer>
        </el-container>
    </div>
</template>

<script setup>
import Main from './Main.vue'
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