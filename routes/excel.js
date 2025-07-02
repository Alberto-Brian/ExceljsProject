import express from 'express';
const router = express.Router();

import { gerarFechoDeCaixa } from '../src/controllers/gerarFechoDeCaixa.js';
import { gerarPlanilhaBasica } from '../src/controllers/gerarPlanilhaBascia.js';
import { gerarRelatorioVendas } from '../src/controllers/excelController.js';
import { gerarListaProdutos } from '../src/controllers/gerarListaProdutos.js';
import { gerarPlanilhaCustomizada } from '../src/controllers/gerarPlanilhaCustomizada.js';
import { gerarFechoDeTurno } from '../src/controllers/gerarFechoDeTurno.js';

router.get('/basic', gerarPlanilhaBasica);
router.get('/fecho-caixa', gerarFechoDeCaixa);
router.get('/vendas', gerarRelatorioVendas);
router.get('/produtos', gerarListaProdutos);
router.post('/custom', gerarPlanilhaCustomizada);
router.get('/fecho-turno', gerarFechoDeTurno);


export default router;
