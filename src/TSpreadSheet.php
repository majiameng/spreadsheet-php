<?php
/**
 * SpreadSheet
 * @author: JiaMeng <666@majiameng.com>
 */
namespace tinymeng\spreadsheet;

// 目录入口
define('SPREADSHEET_ROOT_PATH', dirname(__DIR__));
use tinymeng\tools\Strings;
use tinymeng\spreadsheet\Connector\GatewayInterface;
/**
 * @method static \tinymeng\spreadsheet\Gateways\Export export(array|null $config) 导出
 * @method static \tinymeng\spreadsheet\Gateways\Import import(array|null $config) 导入
 */
abstract class TSpreadSheet
{

    /**
     * Description:  init
     * @author: JiaMeng <666@majiameng.com>
     * Updater:
     * @param $gateway
     * @param null $config
     * @return mixed
     * @throws \Exception
     */
    protected static function init($gateway, $config=[])
    {
        $gateway = Strings::uFirst($gateway);
        $class = __NAMESPACE__ . '\\Gateways\\' . $gateway;
        if (class_exists($class)) {
            $configFile = SPREADSHEET_ROOT_PATH."/config/TSpreadSheet.php";
            if (!file_exists($configFile)) {
                return false;
            }
            $baseConfig = require $configFile;
            $app = new $class(array_replace_recursive($baseConfig,$config));
            if ($app instanceof GatewayInterface) {
                return $app;
            }
            throw new \Exception("基类 [$gateway] 必须继承抽象类 [GatewayInterface]");
        }
        throw new \Exception("基类 [$gateway] 不存在");
    }

    /**
     * Description:  __callStatic
     * @author: JiaMeng <666@majiameng.com>
     * Updater:
     * @param $gateway
     * @param $config
     * @return mixed
     */
    public static function __callStatic($gateway, $config=[])
    {
        return self::init($gateway, ...$config);
    }

}
