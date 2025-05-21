import { Routes } from '@angular/router';
import { MainComponent } from './components/main/main.component';
import { TestComponent } from './components/test/test.component';

export const routes: Routes = [
  {
    path: '',
    component: MainComponent,
    // component: TestComponent,
  },
];
