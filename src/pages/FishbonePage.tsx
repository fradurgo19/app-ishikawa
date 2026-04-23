import React from 'react';
import { useSearchParams, useNavigate, useLocation, Navigate } from 'react-router-dom';
import { FishboneDiagram } from '../organisms/FishboneDiagram';
import { Button } from '../atoms/Button';
import { ArrowLeft } from 'lucide-react';

function isFromDataTableState(state: unknown): boolean {
  return (
    typeof state === 'object' &&
    state !== null &&
    'fromDataTable' in state &&
    (state as { fromDataTable?: boolean }).fromDataTable === true
  );
}

export const FishbonePage: React.FC = () => {
  const [searchParams] = useSearchParams();
  const navigate = useNavigate();
  const location = useLocation();

  if (!isFromDataTableState(location.state)) {
    return <Navigate to="/data-table" replace />;
  }

  const selectedTipoEquipo = normalizeQueryParam(searchParams.get('tipoEquipo'));
  const selectedBrand = normalizeQueryParam(searchParams.get('brand'));
  const selectedModel = normalizeQueryParam(searchParams.get('model'));
  const selectedProblem = normalizeQueryParam(searchParams.get('problem'));

  return (
    <div className="min-h-screen w-full bg-gray-50">
      <div className="w-full px-4 sm:px-6 lg:px-8 py-8">
        <div className="flex items-center gap-4 mb-8">
          <Button
            variant="ghost"
            icon={ArrowLeft}
            onClick={() => navigate('/data-table')}
          >
            Volver a la tabla de datos
          </Button>
          
          <div>
            <h1 className="text-3xl font-bold text-gray-900">Análisis Ishikawa</h1>
            <div className="flex flex-wrap gap-4 text-sm text-gray-600 mt-2">
              {selectedTipoEquipo && <span>Tipo de equipo: {selectedTipoEquipo}</span>}
              {selectedBrand && <span>Marca: {selectedBrand}</span>}
              {selectedModel && <span>Modelo: {selectedModel}</span>}
              {selectedProblem && <span>Problema: {selectedProblem}</span>}
            </div>
          </div>
        </div>

        <FishboneDiagram
          selectedTipoEquipo={selectedTipoEquipo}
          selectedBrand={selectedBrand}
          selectedModel={selectedModel}
          selectedProblem={selectedProblem}
        />

        <div className="mt-8 bg-white rounded-lg p-6 shadow-md">
          <h2 className="text-lg font-semibold text-gray-900 mb-4">Cómo Usar</h2>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6 text-sm text-gray-600">
            <div>
              <h3 className="font-medium text-gray-900 mb-2">Navegación</h3>
              <ul className="space-y-1">
                <li>
                  • Tras expandir marca, los modelos alternan arriba y abajo; cada modelo arrastra toda su rama al mismo
                  lado (arriba o abajo)
                </li>
                <li>• Diferentes colores representan diferentes tipos de datos</li>
                <li>• Los iconos ayudan a identificar tipos de contenido específicos</li>
                <li>
                  • Recurso y Adjunto: si hay enlace, se abre en una nueva pestaña para verlo; si no, se muestra la
                  pantalla de detalle del registro
                </li>
              </ul>
            </div>
            <div>
              <h3 className="font-medium text-gray-900 mb-2">Estructura de Datos</h3>
              <ul className="space-y-1">
                <li>• Tipo de equipo → Marca → Modelo → Sección → Problema</li>
                <li>• Problema → Tipo de Actividad → Actividad</li>
                <li>• Actividad → Recurso, Adjuntos, Tiempo</li>
              </ul>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

function normalizeQueryParam(value: string | null): string | undefined {
  const normalizedValue = value?.trim();
  return normalizedValue || undefined;
}