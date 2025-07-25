<?php

declare(strict_types=1);

namespace Exsol\UltimateNativeXlsbParser\Exception;

/**
 * XLSB Parser Exception
 * 
 * Thrown when errors occur during XLSB parsing operations
 * 
 * @package Exsol\UltimateNativeXlsbParser\Exception
 */
class XlsbParserException extends \Exception
{
    /**
     * @param string $message
     * @param int $code
     * @param \Throwable|null $previous
     */
    public function __construct(string $message = "", int $code = 0, ?\Throwable $previous = null)
    {
        parent::__construct($message, $code, $previous);
    }
}