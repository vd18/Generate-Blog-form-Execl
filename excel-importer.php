<?php
/*
Plugin Name: Excel Importer
Description: Import blog titles and content from an Excel file and generate posts.
Version: 1.0
*/

// Exit if accessed directly
if (!defined('ABSPATH')) {
    exit;
}

// Add menu item to dashboard
add_action('admin_menu', 'excel_importer_add_menu');

function excel_importer_add_menu() {
    add_menu_page('Excel Importer', 'Excel Importer', 'manage_options', 'excel-importer', 'excel_importer_page');
}

// Display plugin page
function excel_importer_page() {
    $msg = isset($_GET['msg']) ? $_GET['msg'] : ''; // Get success message from URL
    ?>
    <div class="wrap">
        <h1>Generate Blogs From Excel</h1>
        <?php if (!empty($msg)): ?>
            <div class="updated">
                <p><?php echo esc_html($msg); ?></p>
            </div>
        <?php endif; ?>
        <form method="post" enctype="multipart/form-data" action="<?php echo esc_url(admin_url('admin-post.php')); ?>">
            <input type="file" name="excel_file">
            <?php wp_nonce_field('excel_import_nonce', 'excel_import_nonce_field'); ?>
            <input type="hidden" name="action" value="excel_import">
            <input type="submit" name="submit" value="Import">
        </form>
    </div>
    <?php
}

// Handle form submission
add_action('admin_post_excel_import', 'excel_importer_handle_form');

function excel_importer_handle_form() {
    $success_msg = ''; // Initialize success message variable

    // Check if nonce is valid
    if (isset($_POST['excel_import_nonce_field']) && wp_verify_nonce($_POST['excel_import_nonce_field'], 'excel_import_nonce')) {
        // Check if file is uploaded
        if (isset($_FILES['excel_file']) && !empty($_FILES['excel_file']['tmp_name'])) {
            // Parse Excel file
            $excel_data = excel_importer_parse_excel($_FILES['excel_file']['tmp_name']);
            // Process parsed data (create posts)
            foreach ($excel_data as $row) {
                $post_id = wp_insert_post(array(
                    'post_title' => $row[0], // First column contains blog title
                    'post_content' => $row[1], // Second column contains blog content
                    'post_status' => 'publish',
                    'post_type' => 'post',
                ));
                if (is_wp_error($post_id)) {
                    // Handle error
                } else {
                    // Post created successfully
                    $success_msg = 'Posts generated successfully.';
                }
            }
        } else {
            // No file uploaded, display error message
            $success_msg = 'No file uploaded.';
        }
    } else {
        // Invalid nonce, display error message
        $success_msg = 'Invalid nonce.';
    }

    // Redirect back to plugin page with success message
    wp_redirect(admin_url('admin.php?page=excel-importer&msg=' . urlencode($success_msg)));
    exit;
}

// Parse Excel file and return data
function excel_importer_parse_excel($file_path) {
    $excel_data = array();
    require_once plugin_dir_path(__FILE__) . 'vendor/autoload.php'; // Include PHPSpreadsheet library

    $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx'); // Specify Excel file format (Xlsx for Excel 2007 and later)
    $spreadsheet = $reader->load($file_path);
    $worksheet = $spreadsheet->getActiveSheet();

    foreach ($worksheet->getRowIterator() as $row) {
        $cellIterator = $row->getCellIterator();
        $rowData = array();
        foreach ($cellIterator as $cell) {
            $rowData[] = $cell->getValue();
        }
        $excel_data[] = $rowData;
    }
    return $excel_data;
}
