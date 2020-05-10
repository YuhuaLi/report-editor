import { AppComponent } from './app.component';
import { NgModule } from '@angular/core';
import { Routes, RouterModule } from '@angular/router';

const routes: Routes = [
  {
    path: '',
    component: AppComponent,
    children: [
      {
        path: 'editor',
        loadChildren: () =>
          import('./editor/editor.module').then((m) => m.EditorModule),
      },
      {
        path: '',
        pathMatch: 'full',
        redirectTo: 'editor',
      },
    ],
  },
];

@NgModule({
  imports: [RouterModule.forRoot(routes)],
  exports: [RouterModule],
})
export class AppRoutingModule {}
