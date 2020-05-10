import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';

import { EditorRoutingModule } from './editor-routing.module';
import { EditorComponent } from './editor.component';
import { EditorPanelComponent } from './editor-panel/editor-panel.component';

@NgModule({
  declarations: [EditorComponent, EditorPanelComponent],
  imports: [CommonModule, EditorRoutingModule],
})
export class EditorModule {}
