<?php

/**
 * The plugin bootstrap file
 *
 * This file is read by WordPress to generate the plugin information in the plugin
 * admin area. This file also includes all of the dependencies used by the plugin,
 * registers the activation and deactivation functions, and defines a function
 * that starts the plugin.
 *
 *
 * @wordpress-plugin
 * Plugin Name:       WC Custom Order Export
 * Plugin URI:        #
 * Description:       This plugin is export orders data.
 * Version:           1.0.0
 * Author:            Ashwini Dubey
 * Author URI:        https://www.upwork.com/freelancers/~019cd2a22666badcb6
 * License:           GPL-2.0+
 * License URI:       http://www.gnu.org/licenses/gpl-2.0.txt
 * Text Domain:       wc-custom-order-export
 * Domain Path:       /languages
 */
defined('ABSPATH') || exit;

define('PLUGIN_WCE_VERSION', '1.0.5');

function wce_enqueue_scripts()
{

    wp_enqueue_style('wc-order-export', plugin_dir_url(__FILE__) . 'css/wc-order-export.css', array(), '1.0.0', 'all');
}
add_action('wp_enqueue_scripts', 'wce_enqueue_scripts');

add_action('admin_enqueue_scripts', 'wce_add_color_picker');
function wce_add_color_picker($hook)
{

    if (is_admin()) {
        wp_enqueue_script('jquery-ui-datepicker');
    }
}


add_action('admin_menu', 'wce_plugin_menu');

function wce_plugin_menu()
{
    add_submenu_page('woocommerce', 'Woo Order Export', 'Woo Order Export', 'manage_options', 'woo-order-export', 'wc_order_export_page');
}

function wc_order_export_page()
{
    include_once 'views/wce-settings.php';
}
// function get_order_count_by($product_id)
// {
//     global $wpdb;

//     return $wpdb->get_var($wpdb->prepare("
//         SELECT DISTINCT count(o.ID)
//         FROM {$wpdb->prefix}posts o
//         INNER JOIN {$wpdb->prefix}woocommerce_order_items oi
//             ON o.ID = oi.order_id
//         INNER JOIN {$wpdb->prefix}woocommerce_order_itemmeta oim
//             ON oi.order_item_id = oim.order_item_id
//         WHERE oim.meta_key IN ('_product_id','_variation_id')
//             AND oim.meta_value = %d
//     ",  $product_id));
// }
