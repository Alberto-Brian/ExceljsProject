import express, { Router } from 'express';
const router: Router = express.Router();

import { gerarFechoDeCaixa } from '../src/controllers/gerarFechoDeCaixa';
import { gerarPlanilhaBasica } from '../src/controllers/gerarPlanilhaBascia.ts';
import { gerarRelatorioVendas } from '../src/controllers/excelController.ts';
import { gerarListaProdutos } from '../src/controllers/gerarListaProdutos.ts';
import { gerarPlanilhaCustomizada } from '../src/controllers/gerarPlanilhaCustomizada.ts';
import { gerarFechoDeTurno } from '../src/controllers/gerarFechoDeTurno.ts';

router.get('/basic', gerarPlanilhaBasica);
router.get('/fecho-caixa', gerarFechoDeCaixa);
router.get('/vendas', gerarRelatorioVendas);
router.get('/produtos', gerarListaProdutos);
router.post('/custom', gerarPlanilhaCustomizada);
router.get('/fecho-turno', gerarFechoDeTurno);

export default router;