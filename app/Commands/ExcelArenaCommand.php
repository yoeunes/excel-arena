<?php

namespace App\Commands;

use Carbon\Carbon;
use LaravelZero\Framework\Commands\Command;
use Maatwebsite\Excel\Classes\LaravelExcelWorksheet;
use Maatwebsite\Excel\Collections\CellCollection;
use Maatwebsite\Excel\Excel;
use Maatwebsite\Excel\Readers\LaravelExcelReader;
use Maatwebsite\Excel\Writers\LaravelExcelWriter;

class ExcelArenaCommand extends Command
{
    protected const DEPART    = 1;
    protected const COMPAGNIE = 2;
    protected const NUM_VOL   = 3;
    protected const SIEGE     = 12;

    /**
     * The name and signature of the command.
     *
     * @var string
     */
    protected $signature = 'format {filename}';

    /**
     * The description of the command.
     *
     * @var string
     */
    protected $description = 'Format an excel file to be read inside arena.';

    protected $formatedData = [];

    protected $skipRow = true;

    /** @var Excel */
    protected $excel;

    protected $pwd = '';

    public function __construct(Excel $excel)
    {
        parent::__construct();

        $this->excel = $excel;
        $this->pwd = getcwd() . '/';
    }

    /**
     * Execute the command. Here goes the code.
     *
     * @return void
     */
    public function handle(): void
    {
        if(empty($filename = $this->argument('filename')) || !file_exists($this->pwd . $filename)) {
            $this->error('le fichier ' . $filename . ' n\'existe pas.');

            return;
        }

        $this->excel->load($filename, function(LaravelExcelReader $reader) {
            $reader
                ->noHeading()
                ->ignoreEmpty()
                ->formatDates(false);

            $reader->each(function(CellCollection $cell) {
                $data = $cell->toArray();

                if('ARR' === $data[0]) {
                    $this->skipRow = false;
                    return;
                }

                if(true === $this->skipRow) {
                    return;
                }

                if(empty($data[self::DEPART])) {
                    return;
                }

                $siege = (int) $data[self::SIEGE];
                $thirdSiege = (int) ($siege * 0.3);

                $this->formatedData[] = [
                    'depart'   => Carbon::createFromFormat('H:i', $data[self::DEPART])->subHours(1)->format('H:i'),
                    'companie' => $data[self::COMPAGNIE],
                    'num_vol'  => $data[self::NUM_VOL],
                    'siege'    => $thirdSiege,
                ];

                $this->formatedData[] = [
                    'depart'   => Carbon::createFromFormat('H:i', $data[self::DEPART])->subHours(2)->format('H:i'),
                    'companie' => $data[self::COMPAGNIE],
                    'num_vol'  => $data[self::NUM_VOL],
                    'siege'    => $siege - 2 * $thirdSiege,
                ];

                $this->formatedData[] = [
                    'depart'   => Carbon::createFromFormat('H:i', $data[self::DEPART])->subHours(3)->format('H:i'),
                    'companie' => $data[self::COMPAGNIE],
                    'num_vol'  => $data[self::NUM_VOL],
                    'siege'    => $thirdSiege,
                ];
            });
        });
        
        $this->formatedData = collect($this->formatedData)->sortBy(function($item) {
            return Carbon::createFromFormat('H:i', $item['depart']);
        })->values()->toArray();

        foreach ($this->formatedData as $index => $item) {
            $lastDate = 0 === $index ? Carbon::now()->startOfDay() : Carbon::createFromFormat('H:i', $this->formatedData[$index - 1]['depart']);
            $diffInMinutes = Carbon::createFromFormat('H:i', $item['depart'])->diffInMinutes($lastDate);
            $this->formatedData[$index]['temps_attente'] = $diffInMinutes;
        }

        $content = $this->excel->create($this->pwd . str_replace('in', 'out', $filename), function(LaravelExcelWriter $writer) {
            $writer->sheet('LUNDI', function(LaravelExcelWorksheet $worksheet) {
                $worksheet->fromArray($this->formatedData, null, 'A1', true, false);
            });
        })->string('xls');

        file_put_contents($this->pwd . str_replace('in', 'out', $filename), $content);
    }
}
