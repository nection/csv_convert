<?php

namespace Drupal\csv_convert\Controller;

use Drupal\Core\Controller\ControllerBase;
use Drupal\Core\Database\Connection; // Afegit per la injecció de dependències BD
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\HttpFoundation\StreamedResponse; // Afegit per a streaming eficient
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Symfony\Component\DependencyInjection\ContainerInterface; // Afegit per la injecció de dependències

/**
 * Controlador per gestionar la descàrrega de dades del formulari en format CSV i Excel.
 * Llegeix les dades directament de la base de dades.
 */
class CsvController extends ControllerBase {

  /**
   * The database connection.
   *
   * @var \Drupal\Core\Database\Connection
   */
  protected $database; // Propietat per guardar la connexió a la BD

  /**
   * Constructs a new CsvController object.
   *
   * @param \Drupal\Core\Database\Connection $database
   *   The database connection.
   */
  public function __construct(Connection $database) { // Constructor per injectar la BD
    $this->database = $database;
  }

  /**
   * {@inheritdoc}
   */
  public static function create(ContainerInterface $container) { // Mètode estàtic per a la injecció de dependències
    return new static(
      $container->get('database')
    );
  }

  /**
   * Mostra la pàgina amb els botons per gestionar la conversió.
   * Accessible només per als rols 'administrator' i 'gestor'.
   */
  public function showPage() {
    $current_user = \Drupal::currentUser();
    $es_gestor = $current_user->hasRole('gestor');
    // Incloem 'administrator' aquí també per simplificar la comprovació
    $es_administrador = $current_user->hasRole('administrator') || $es_gestor;

    // Comprova si l'usuari té el rol adequat
    if (!$es_administrador) {
      // Retornem una resposta prohibida, més adequat que només text.
       return new Response($this->t('No tens permisos per accedir a aquesta pàgina.'), 403);
    }

    // Obtenim el base path del lloc per assegurar que les rutes inclouen el subdirectori si n'hi ha
    $base_path = \Drupal::request()->getBasePath();

    // Construïm les URLs completes
    // Nota: Encara que Url::fromRoute és més 'Drupal way', mantenim la concatenació
    //       original per minimitzar canvis fora de l'objectiu principal.
    $url_excel = $base_path . '/csv/download-excel';
    $url_csv = $base_path . '/csv/download-csv';

    // Contingut de la pàgina amb els botons.
    $output = '
      <h1>' . $this->t('Gestió de fitxers de dades del formulari') . '</h1>
      <p>' . $this->t('Aquesta pàgina permet descarregar les dades dels formularis enviats, llegides directament des de la base de dades, en format Excel o CSV.') . '</p>
      <a href="' . $url_excel . '" class="button button--primary">' . $this->t('Descarregar Excel (.xlsx)') . '</a><br><br>
      <a href="' . $url_csv . '" class="button button--primary">' . $this->t('Descarregar CSV') . '</a><br><br>
    ';

    return [
      '#markup' => $output,
    ];
  }

  /**
   * Genera una fulla de càlcul Excel des de la base de dades i permet la descàrrega.
   * Només accessible per als rols 'administrator' i 'gestor'.
   */
  public function downloadExcel() {
    $current_user = \Drupal::currentUser();
    $es_gestor = $current_user->hasRole('gestor');
    $es_administrador = $current_user->hasRole('administrator') || $es_gestor;

    if (!$es_administrador) {
      // Retornem una resposta prohibida si no té permisos.
       return new Response($this->t('No tens permisos per descarregar aquest fitxer.'), 403);
    }

    // Crear l'objecte Spreadsheet
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    $filename = "dades_equipaments_" . date('Ymd_His') . ".xlsx";

    // Utilitza StreamedResponse per a fitxers grans
    $response = new StreamedResponse(function() use ($sheet, $spreadsheet) {
      try {
        // Consultar les dades de la base de dades
        $query = $this->database->select('nou_formulari_dades_formulari', 't');
        $query->fields('t'); // Seleccionar totes les columnes
        // Opcional: afegir orderBy si vols un ordre específic
        // $query->orderBy('hora_submissio', 'DESC');
        $results = $query->execute();

        $headers = [];
        $row_index = 1; // Comença a la fila 1 per a les capçaleres

        foreach ($results as $record) {
          // Convertir objecte StdClass a array associatiu
          $record_array = (array) $record;

          // A la primera iteració, obtenir i escriure les capçaleres
          if ($row_index === 1) {
            $headers = array_keys($record_array);
            $col_index = 1;
            foreach ($headers as $header) {
              // Escriu la capçalera
              $sheet->setCellValueByColumnAndRow($col_index, $row_index, $header);
              // Opcional: Autoajustar amplada columna
              $sheet->getColumnDimensionByColumn($col_index)->setAutoSize(TRUE);
              $col_index++;
            }
            $row_index++; // Passar a la següent fila per a les dades
          }

          // Escriure les dades de la fila actual assegurant l'ordre de les capçaleres
          $col_index = 1;
          foreach ($headers as $header) { // Iterar per l'ordre de les capçaleres
            // Utilitza l'operador null coalescing per si alguna columna no existís (poc probable aquí)
             $value = $record_array[$header] ?? '';
             // Tractament especial per a valors que semblen números però són text (ex: codi equipament)
             if (is_string($value) && !is_numeric($value) && preg_match('/^\d+$/', $value)) {
                $sheet->getCellByColumnAndRow($col_index, $row_index)
                      ->setValueExplicit($value, \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
             } else {
                $sheet->setCellValueByColumnAndRow($col_index, $row_index, $value);
             }
             $col_index++;
          }
          $row_index++;
        }

        // Si no hi ha resultats, el bucle no s'executarà.
        // Podries afegir un missatge si la taula està buida.
        if ($row_index === 1) {
           $sheet->setCellValueByColumnAndRow(1, 1, $this->t('No hi ha dades disponibles a la taula nou_formulari_dades_formulari.'));
        }

        // Crear l'arxiu Excel per a la descàrrega
        $writer = new Xlsx($spreadsheet);
        $writer->save('php://output');

      } catch (\Exception $e) {
        $this->getLogger('csv_convert')->error('Error generant Excel des de la BD: @message', ['@message' => $e->getMessage()]);
        // Escriu un missatge d'error directament a la sortida si falla durant el streaming
        echo $this->t("S'ha produït un error en generar el fitxer Excel. Si us plau, revisa els logs del sistema.");
      }
    });

    $response->headers->set('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    $response->headers->set('Content-Disposition', 'attachment;filename="' . $filename . '"');
    $response->headers->set('Cache-Control', 'max-age=0');
    $response->headers->set('Pragma', 'public'); // Afegit per compatibilitat

    return $response;
  }

  /**
   * Genera un fitxer CSV des de la base de dades i permet la descàrrega.
   * Només accessible per als rols 'administrator' i 'gestor'.
   */
  public function downloadCsv() {
    $current_user = \Drupal::currentUser();
    $es_gestor = $current_user->hasRole('gestor');
    $es_administrador = $current_user->hasRole('administrator') || $es_gestor;

    if (!$es_administrador) {
      // Retornem una resposta prohibida si no té permisos.
       return new Response($this->t('No tens permisos per descarregar aquest fitxer.'), 403);
    }

    $filename = "dades_formulari_" . date('Ymd_His') . ".csv";

    // Utilitza StreamedResponse per gestionar la sortida CSV eficientment
    $response = new StreamedResponse(function() {
        try {
            // Consultar les dades de la base de dades
            $query = $this->database->select('nou_formulari_dades_formulari', 't');
            $query->fields('t'); // Seleccionar totes les columnes
            // Opcional: afegir orderBy si vols un ordre específic
            // $query->orderBy('hora_submissio', 'DESC');
            $results = $query->execute();

            // Obrir el flux de sortida per escriure el CSV
            $handle = fopen('php://output', 'w');
            if ($handle === FALSE) {
                 // Llança excepció si no es pot obrir el flux
                 throw new \RuntimeException("No s'ha pogut obrir php://output per escriure el CSV.");
            }

            // Afegir BOM (Byte Order Mark) per a UTF-8 per millorar compatibilitat amb Excel
            fwrite($handle, "\xEF\xBB\xBF");

            $headers = [];
            $first_row = TRUE;

            // Iterar sobre els resultats
            foreach ($results as $record) {
                // Convertir objecte StdClass a array associatiu
                $record_array = (array) $record;

                // A la primera fila, obtenir i escriure les capçaleres
                if ($first_row) {
                    $headers = array_keys($record_array);
                    fputcsv($handle, $headers); // Escriu la línia de capçalera
                    $first_row = FALSE;
                }

                // Ordenar les dades segons les capçaleres per assegurar l'ordre correcte
                $data_row = [];
                foreach($headers as $header) {
                    // Afegeix el valor o un string buit si no existeix la clau (poc probable)
                    $data_row[] = $record_array[$header] ?? '';
                }

                // Escriure la fila de dades al CSV
                fputcsv($handle, $data_row);
            }

            // Si no hi ha resultats ($first_row encara és TRUE), el fitxer estarà buit (excepte el BOM).
            // Podries escriure només les capçaleres si les coneguessis i la taula estigués buida.
            // if ($first_row) {
                 // $default_headers = ['id', 'username', 'comarca', ...]; // Defineix les capçaleres esperades
                 // fputcsv($handle, $default_headers);
            // }

            // Tancar el gestor de fitxers
            fclose($handle);

        } catch (\Exception $e) {
            $this->getLogger('csv_convert')->error('Error generant CSV des de la BD: @message', ['@message' => $e->getMessage()]);
            // Escriu un missatge d'error directament a la sortida si falla durant el streaming
            echo $this->t("S'ha produït un error en generar el fitxer CSV. Si us plau, revisa els logs del sistema.");
            // Tanquem el handle si encara està obert en cas d'error
             if (isset($handle) && is_resource($handle)) {
                 fclose($handle);
             }
        }
    });

    // Configurar les capçaleres de la resposta HTTP per a CSV
    $response->headers->set('Content-Type', 'text/csv; charset=utf-8'); // Especifica UTF-8
    $response->headers->set('Content-Disposition', 'attachment;filename="' . $filename . '"');
    // Capçaleres addicionals per evitar problemes de cache
    $response->headers->set('Cache-Control','private, no-cache, must-revalidate');
    $response->headers->set('Pragma', 'private');
    $response->headers->set('Expires', '0'); // Expira immediatament

    return $response;
  }
}